// ============================================================
// PRIMERS CONTROL — backend de administração de usuários
// Roda como Supabase Edge Function. Guarda a chave secreta
// (service_role) do lado do servidor — ela NUNCA vai para o
// código do site. Só quem estiver logado como "admin" consegue
// usar estas ações.
//
// Ações suportadas (enviadas no corpo da requisição em { action }):
//   "list"   -> lista os usuários cadastrados (login, nome, papel)
//   "create" -> cria um novo acesso (login, senha, nome, papel)
//   "delete" -> remove um acesso existente (id do usuário)
// ============================================================
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;

const admin = createClient(SUPABASE_URL, SERVICE_ROLE_KEY, {
  auth: { autoRefreshToken: false, persistSession: false },
});

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

function jsonResponse(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  try {
    // 1) Confirma quem está chamando (token do usuário logado, enviado pelo app)
    const authHeader = req.headers.get("Authorization") ?? "";
    const token = authHeader.replace("Bearer ", "");
    if (!token) return jsonResponse({ error: "Não autenticado." }, 401);

    const { data: userData, error: userErr } = await admin.auth.getUser(token);
    if (userErr || !userData?.user) return jsonResponse({ error: "Sessão inválida." }, 401);

    // 2) Confirma que quem está chamando é administrador
    const { data: perfil, error: perfilErr } = await admin
      .from("profiles")
      .select("role")
      .eq("id", userData.user.id)
      .single();
    if (perfilErr || perfil?.role !== "admin") {
      return jsonResponse({ error: "Apenas administradores podem gerenciar acessos." }, 403);
    }

    const body = await req.json();
    const { action } = body;

    // ---------------- LISTAR ----------------
    if (action === "list") {
      const { data: perfis, error } = await admin.from("profiles").select("id, usuario, nome, role, created_at").order("created_at");
      if (error) return jsonResponse({ error: error.message }, 400);
      return jsonResponse({ usuarios: perfis });
    }

    // ---------------- CRIAR ----------------
    if (action === "create") {
      const { usuario, senha, nome, role } = body;
      if (!usuario || !senha || !nome || !role) {
        return jsonResponse({ error: "Preencha usuário, senha, nome e papel." }, 400);
      }
      if (senha.length < 6) return jsonResponse({ error: "A senha precisa ter ao menos 6 caracteres." }, 400);

      const emailInterno = `${usuario.trim().toLowerCase()}@primers.local`;
      const { data: novoUsuario, error: createErr } = await admin.auth.admin.createUser({
        email: emailInterno,
        password: senha,
        email_confirm: true,
      });
      if (createErr) return jsonResponse({ error: createErr.message }, 400);

      const { error: profileErr } = await admin.from("profiles").insert({
        id: novoUsuario.user.id,
        usuario: usuario.trim().toLowerCase(),
        nome,
        role,
      });
      if (profileErr) {
        await admin.auth.admin.deleteUser(novoUsuario.user.id); // desfaz se o perfil falhar
        return jsonResponse({ error: profileErr.message }, 400);
      }
      return jsonResponse({ ok: true, id: novoUsuario.user.id });
    }

    // ---------------- REMOVER ----------------
    if (action === "delete") {
      const { id } = body;
      if (!id) return jsonResponse({ error: "Informe o id do acesso a remover." }, 400);
      if (id === userData.user.id) return jsonResponse({ error: "Você não pode remover o seu próprio acesso." }, 400);

      const { error: delErr } = await admin.auth.admin.deleteUser(id);
      if (delErr) return jsonResponse({ error: delErr.message }, 400);
      await admin.from("profiles").delete().eq("id", id);
      return jsonResponse({ ok: true });
    }

    return jsonResponse({ error: "Ação desconhecida." }, 400);
  } catch (e) {
    return jsonResponse({ error: String(e) }, 500);
  }
});
