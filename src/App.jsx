import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend
} from "recharts";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import {
  LayoutDashboard, ListChecks, Search, Plus, Upload, Eye, EyeOff, X,
  Building2, AlertTriangle, Clock, ChevronRight, FileStack, Timer,
  ClipboardCheck, History, Download, MessageSquarePlus, CheckCircle2,
  XCircle, MinusCircle, Filter, ChevronLeft, ChevronDown, DollarSign,
  Link2, Lock, Wrench, FileSignature, Pencil, Trash2, PlusCircle, LogOut, Users
} from "lucide-react";

/* ============================================================
   CONEXÃO COM O BANCO DE DADOS (Supabase)
   Troque os dois valores abaixo pelos do SEU projeto Supabase:
   painel do Supabase → Project Settings → API → "Project URL" e
   a chave "anon public" (a chave "service_role" NUNCA vai aqui).
   ============================================================ */
const SUPABASE_URL = "COLE_AQUI_A_URL_DO_SEU_PROJETO_SUPABASE";
const SUPABASE_ANON_KEY = "COLE_AQUI_A_CHAVE_ANON_PUBLIC";
const SUPABASE_CONFIGURADO = SUPABASE_URL.startsWith("http");
const supabase = createClient(
  SUPABASE_CONFIGURADO ? SUPABASE_URL : "https://placeholder.supabase.co",
  SUPABASE_CONFIGURADO ? SUPABASE_ANON_KEY : "placeholder-key-not-configured"
);

/* ============================================================
   DESIGN TOKENS — identidade PRIMERS
   ============================================================ */
const COLORS = {
  bg: "#0a1420", panel: "#101f30", panelAlt: "#16283d", panelSoft: "#132436",
  border: "rgba(255,255,255,0.08)", borderStrong: "rgba(255,255,255,0.14)",
  steel: "#8493a6", steelLight: "#b7c2cf", ice: "#eef2f6",
  red: "#e1483d", redDim: "rgba(225,72,61,0.16)",
  green: "#3ecf8e", greenDim: "rgba(62,207,142,0.15)",
  yellow: "#e8c547", yellowDim: "rgba(232,197,71,0.15)",
  orange: "#f2894b", orangeDim: "rgba(242,137,75,0.15)",
  blue: "#4a90d9", blueDim: "rgba(74,144,217,0.15)",
  gray: "#5b6675", grayDim: "rgba(91,102,117,0.18)",
  overdue: "#c23b32", overdueDim: "rgba(194,59,50,0.20)",
};

/* ============================================================
   MODELO DE STATUS — cada status tem um responsável pela próxima ação
   ============================================================ */
const STATUS_CONFIG = {
  aguardando: { label: "Aguardando", responsavel: "Primers", fg: COLORS.steel, bg: COLORS.grayDim, final: false, grupo: "aguardando" },
  iniciado: { label: "Iniciado", responsavel: "Primers", fg: COLORS.blue, bg: COLORS.blueDim, final: false, grupo: "aguardando" },
  em_montagem: { label: "Em montagem", responsavel: "Primers", fg: COLORS.orange, bg: COLORS.orangeDim, final: false, grupo: "montagem" },
  protocolado: { label: "Protocolado — aguardando análise", responsavel: "Órgão", fg: COLORS.blue, bg: COLORS.blueDim, final: false, grupo: "analise" },
  exigencia_primers: { label: "Em exigência — aguardando atendimento", responsavel: "Primers", fg: COLORS.orange, bg: COLORS.orangeDim, final: false, grupo: "exigencia" },
  exigencia_cliente: { label: "Em exigência — aguardando retorno do cliente", responsavel: "Cliente", fg: COLORS.yellow, bg: COLORS.yellowDim, final: false, grupo: "exigencia" },
  aguardando_orgao: { label: "Aguardando retorno do órgão", responsavel: "Órgão", fg: COLORS.blue, bg: COLORS.blueDim, final: false, grupo: "analise" },
  concluido: { label: "Concluído / Deferido", responsavel: "Finalizado", fg: COLORS.green, bg: COLORS.greenDim, final: true, grupo: "concluido" },
  indeferido: { label: "Indeferido", responsavel: "Finalizado", fg: "#ffb3ac", bg: COLORS.overdueDim, final: true, grupo: "indeferido" },
  cancelado: { label: "Cancelado", responsavel: "Finalizado", fg: COLORS.steel, bg: COLORS.grayDim, final: true, grupo: "cancelado" },
  suspenso: { label: "Suspenso", responsavel: "Cliente", fg: COLORS.yellow, bg: COLORS.yellowDim, final: false, grupo: "exigencia" },
};
const STATUS_KEYS = Object.keys(STATUS_CONFIG);
const RESPONSAVEIS = ["Primers", "Cliente", "Órgão", "Finalizado"];

const ATUALIZACAO_TIPOS = [
  "Atendimento presencial", "Atendimento online", "Comunique-se / Exigência recebida",
  "Reunião / Alinhamento com cliente", "Tramitação / Movimentação processual",
  "Cobrança de celeridade ao órgão", "Retorno do cliente", "Outro",
];

/* ============================================================
   TEMPLATES DE CHECKLIST TÉCNICO
   ============================================================ */
const CHECKLIST_TEMPLATES = {
  default: {
    nome: "Checklist padrão de serviço",
    secoes: [
      { titulo: "Preparação", itens: [
        { texto: "Documentação societária / cadeia de poderes conferida" },
        { texto: "Procuração e ART/RRT do responsável técnico" },
        { texto: "Documentos do imóvel (matrícula, IPTU) conferidos" },
        { texto: "Requisitos específicos do órgão levantados" },
        { texto: "Pré-protocolo / análise técnica concluída" },
      ]},
      { titulo: "Protocolo", itens: [
        { texto: "Taxas e emolumentos pagos" },
        { texto: "Documentação do cliente completa e validada" },
        { texto: "Conferência final antes do protocolo" },
      ]},
    ],
  },
  pmsp_projeto: {
    nome: "Checklist — Aprovação de Projeto (PMSP)",
    secoes: [
      { titulo: "Levantamento Planialtimétrico", itens: [
        { texto: "Medidas do terreno R=Real / E=Escritura" },
        { texto: "Medidas lineares do perímetro (divergência R x E ≤ 5%)" },
        { texto: "Área total do terreno R e E (divergência ≤ 5%)" },
        { texto: "Ângulos internos do terreno" },
        { texto: "Cotas de nível (vértices, área interna, calçada e rua) — compatibilizar com Geosampa" },
        { texto: "Edificação existente (a reformar / demolir / regularizar)" },
        { texto: "Árvores internas existentes" },
        { texto: "Área permeável existente" },
        { texto: "Largura de calçada e via (até o outro lado da rua)" },
        { texto: "Nº da matrícula indicado em planta" },
        { texto: "CODLOG da rua indicado em planta" },
        { texto: "Nº do SQL indicado em planta" },
        { texto: "Numeração predial e SQL dos vizinhos" },
        { texto: "Infraestrutura existente indicada (árvores, postes, semáforos, boca de lobo)" },
        { texto: "Curvas de nível (base Geosampa)" },
        { texto: "Legenda (existente regular/irregular, a reformar/regularizar/demolir)" },
        { texto: "Notas de projeto obrigatórias", detalhe: "Córregos/águas/galerias; árvores no local; postes e mobiliário urbano em frente ao lote; demolição (m² e hachura); rede de gás; rede de água/esgoto; passeio público (Decreto 58.611/2019)." },
        { texto: "Carimbo completo", detalhe: "Título, assunto, endereço + CEP, CODLOG, Prefeitura Regional, categoria de uso, zoneamento, quota ambiental, proprietário (CPF/CNPJ), SQL, escala, quadro de áreas (R e E + matrícula), declarações e assinaturas (proprietário, autor do projeto, responsável técnico com CREA/CAU/CCM)." },
      ]},
      { titulo: "Plantas", itens: [
        { texto: "Áreas permeáveis propostas indicadas e numeradas" },
        { texto: "Vagas PCD / Idoso / Carga-Descarga / Motocicletas indicadas" },
        { texto: "Sentido de fluxo interno de veículos" },
        { texto: "Entrada e saída de veículos e pedestres" },
        { texto: "Passeio público cimentado em toda a extensão" },
        { texto: "Rebaixos de guia propostos" },
        { texto: "Áreas computáveis e não computáveis corretas" },
        { texto: "Edificações numeradas, categoria de uso, hachuras, área e cota de nível" },
        { texto: "Perspectiva simplificada dos reservatórios (se Quota Ambiental)" },
        { texto: "Pavimentos superiores sem repetir dados de rua/calçada" },
        { texto: "Linha de corte até o limite do terreno e da rua" },
        { texto: "Quadro de áreas por pavimento + total geral" },
        { texto: "Quadro de áreas: TO, CA e quota de garagem utilizada" },
        { texto: "Quadro de vagas de automóveis conforme padrão da Prefeitura" },
        { texto: "Quadro de áreas permeáveis" },
        { texto: "Legenda completa conforme o projeto pretendido" },
        { texto: "Notas obrigatórias de legislação", detalhe: "Leis 16.050/2014, 16.402/2016, 16.642/2017 + Decreto 57.776/2017 (águas pluviais, estacionamento, instalações sanitárias), rede de gás/água/esgoto, Decreto 58.611/19, aquecimento solar, lotação e sanitários." },
        { texto: "Notas conforme uso e licenças envolvidas", detalhe: "Dimensões de vagas, rampas/inclinação, distância a sanitários, acessibilidade (NBR 9050), iluminação/ventilação mecânica, vestiários, amianto (Lei 13.113/2011), aberturas para vizinho, aeração/insolação." },
        { texto: "Quadro de uso e ocupação do solo completo", detalhe: "Macrozona, macroárea, zona de uso, TO, CA mínimo/básico/máximo/utilizado, permeabilidade, vagas, cota de garagem, área não computável, área construída total." },
        { texto: "Carimbo (conforme levantamento, alterando título e quadro de áreas)" },
      ]},
      { titulo: "Cortes", itens: [
        { texto: "Cortes transversal e longitudinal por toda a extensão do terreno até a rua" },
        { texto: "PNT — Perfil Natural do Terreno indicado (base Geosampa)" },
        { texto: "Volume de aterro demonstrado por hachuras" },
        { texto: "Muros de arrimo cotados (planta e corte)" },
        { texto: "Limite do terreno indicado" },
        { texto: "Corte apenas das alturas" },
        { texto: "Alturas cotadas a partir do PNT (início e fim da edificação)" },
        { texto: "Gabarito de altura cotado a partir do PNT" },
      ]},
      { titulo: "Planta de Quota Ambiental", itens: [
        { texto: "Edificações hachuradas em cinza (\"edificação / área ocupada\")" },
        { texto: "Itens de pontuação apresentados (permeável, piso drenante, muro/cobertura verde, vegetação)" },
        { texto: "Perspectiva simplificada dos reservatórios (Art. 79 e 80 da Lei 16.402)" },
        { texto: "Quadro de composição da pontuação da Quota Ambiental" },
        { texto: "Carimbo, notas, legenda e quadro conforme as plantas" },
      ]},
    ],
  },
};

function pickTemplateId(assunto) {
  const a = (assunto || "").toLowerCase();
  if (a.includes("aprovação de projeto") || a.includes("aprovacao de projeto")) return "pmsp_projeto";
  return "default";
}

function instantiateChecklist(templateId) {
  const tpl = CHECKLIST_TEMPLATES[templateId] || CHECKLIST_TEMPLATES.default;
  return {
    templateId,
    secoes: tpl.secoes.map((sec, si) => ({
      titulo: sec.titulo,
      itens: sec.itens.map((it, ii) => ({
        id: `${si}-${ii}`, texto: it.texto, detalhe: it.detalhe || null,
        estado: "pendente", obs: "",
      })),
    })),
  };
}

/* ============================================================
   DOCUMENTOS EXIGIDOS — checklist documental por tipo de serviço
   (controle de qualidade do protocolo: o que precisa estar no
   processo, e não apenas o desenho em si)
   ============================================================ */
const DOC_TEMPLATES = {
  default: {
    nome: "Documentos padrão de serviço",
    categorias: [
      { titulo: "Documentação geral", itens: [
        { texto: "Documento de identificação do requerente / representante", obrigatorio: true },
        { texto: "Comprovante de propriedade ou posse do imóvel", obrigatorio: true },
        { texto: "Procuração (quando houver procurador)", obrigatorio: false, condicionalLabel: "Há procurador neste serviço?" },
      ]},
    ],
  },
  pmsp_projeto: {
    nome: "Documentos — Aprovação de Projeto (PMSP)",
    categorias: [
      { titulo: "Proprietário do imóvel", itens: [
        { texto: "Matrícula atualizada do imóvel, registrando o proprietário", obrigatorio: true },
        { texto: "Contrato Social (se proprietário pessoa jurídica)", obrigatorio: false, condicionalLabel: "Proprietário é pessoa jurídica?" },
        { texto: "Ata de Eleição dos representantes (se pessoa jurídica)", obrigatorio: false, condicionalLabel: "Proprietário é pessoa jurídica?" },
        { texto: "Documento de identificação do representante legal", obrigatorio: true },
      ]},
      { titulo: "Procurador", itens: [
        { texto: "Procuração assinada por representante legal devidamente constituído", obrigatorio: false, condicionalLabel: "Há procurador neste serviço?" },
        { texto: "Documento de identificação do procurador", obrigatorio: false, condicionalLabel: "Há procurador neste serviço?" },
        { texto: "Contrato de locação com cláusula de outorga de poderes ao locatário", obrigatorio: false, condicionalLabel: "Requerente é o locatário (imóvel locado)?", detalhe: "Indispensável quando o locatário será registrado como requerente perante o órgão." },
      ]},
      { titulo: "Dados do terreno", itens: [
        { texto: "Conferência de medidas (linear e área) — divergência ≤ 5% entre matrícula e levantamento", obrigatorio: true },
        { texto: "Endereço completo e CEP", obrigatorio: true },
        { texto: "SQL", obrigatorio: true },
        { texto: "CODLOG", obrigatorio: true },
        { texto: "Subprefeitura", obrigatorio: true },
        { texto: "Macroárea", obrigatorio: true },
        { texto: "Macrozona", obrigatorio: true },
        { texto: "Perímetro de Qualificação Ambiental", obrigatorio: true },
        { texto: "Zona de uso", obrigatorio: true },
      ]},
      { titulo: "Declarações", itens: [
        { texto: "Declaração de conformidade quanto aos aspectos interiores da edificação (COE e legislação correlata)", obrigatorio: true },
        { texto: "Declaração ou Inexigibilidade — COMAER", obrigatorio: true },
        { texto: "Declaração de Movimento de Terra", obrigatorio: false, condicionalLabel: "Há movimento de terra no projeto?" },
        { texto: "Declaração para licenciamento de equipamento mecânico (transporte permanente / tanque / bomba / filtro de combustível / sistema especial de segurança)", obrigatorio: false, condicionalLabel: "Há equipamento mecânico permanente, tanque ou sistema especial de segurança?" },
      ]},
    ],
  },
};

function instantiateDocumentos(templateId) {
  const tpl = DOC_TEMPLATES[templateId] || DOC_TEMPLATES.default;
  return {
    templateId,
    categorias: tpl.categorias.map((cat, ci) => ({
      titulo: cat.titulo,
      itens: cat.itens.map((it, ii) => ({
        id: `${ci}-${ii}`, texto: it.texto, obrigatorio: it.obrigatorio, condicionalLabel: it.condicionalLabel || null,
        detalhe: it.detalhe || null, aplicavel: !it.condicionalLabel, estado: "pendente", arquivo: null, obs: "", validade: null,
      })),
    })),
  };
}

function documentosProgress(documentos) {
  let total = 0, ok = 0;
  documentos.categorias.forEach((c) => c.itens.forEach((it) => {
    if (!it.aplicavel) return;
    total += 1;
    if (it.estado === "anexado" || it.estado === "nao_aplica") ok += 1;
  }));
  return total === 0 ? 100 : Math.round((ok / total) * 100);
}

function emptyParametrosUrbanisticos() {
  return {
    macrozona: "", macroarea: "", zonaUso: "", perimetroQualificacaoAmbiental: "",
    areaTerrenoEscritura: "", areaTerrenoReal: "", doacaoPasseio: "", areaRemanescente: "",
    taxaOcupacaoMaximaZona: "", areaProjecaoMaxima: "", taxaOcupacaoProjeto: "", areaProjecaoProjeto: "",
    caBasico: "", caMaximo: "", caUtilizado: "", areaComputavelUtilizada: "",
    taxaPermeabilidadeMinima: "", areaPermeavelExigida: "", taxaPermeabilidadeAdotada: "", areaPermeavelAdotada: "",
    areaNaoComputavel: "", areaConstruidaTotal: "",
  };
}

const PARAM_URB_FIELDS = [
  ["macrozona", "Macrozona"], ["macroarea", "Macroárea"], ["zonaUso", "Zona de uso"],
  ["perimetroQualificacaoAmbiental", "Perímetro de Qualificação Ambiental"],
  ["areaTerrenoEscritura", "Área de terreno — Escritura (m²)"], ["areaTerrenoReal", "Área de terreno — Real (m²)"],
  ["doacaoPasseio", "Doação de passeio público (m²)"], ["areaRemanescente", "Área remanescente (m²)"],
  ["taxaOcupacaoMaximaZona", "Taxa de ocupação máxima da zona"], ["areaProjecaoMaxima", "Área de projeção máxima (m²)"],
  ["taxaOcupacaoProjeto", "Taxa de ocupação utilizada no projeto"], ["areaProjecaoProjeto", "Área de projeção do projeto (m²)"],
  ["caBasico", "Coeficiente de aproveitamento básico"], ["caMaximo", "Coeficiente de aproveitamento máximo"],
  ["caUtilizado", "Coeficiente de aproveitamento utilizado"], ["areaComputavelUtilizada", "Área computável utilizada (m²)"],
  ["taxaPermeabilidadeMinima", "Taxa de permeabilidade mínima"], ["areaPermeavelExigida", "Área permeável exigida (m²)"],
  ["taxaPermeabilidadeAdotada", "Taxa de permeabilidade adotada"], ["areaPermeavelAdotada", "Área permeável adotada (m²)"],
  ["areaNaoComputavel", "Área não computável (m²)"], ["areaConstruidaTotal", "Área construída total da edificação (m²)"],
];

function emptyEnquadramentos() {
  return {
    alargamentoPasseio: { aplica: false, obs: "" },
    doacaoDUP: { aplica: false, obs: "" },
    areaTombada: { aplica: false, obs: "" },
    anuenciaMetro: { aplica: false, obs: "" },
  };
}
const ENQUADRAMENTOS_LABELS = {
  alargamentoPasseio: "Alargamento de passeio",
  doacaoDUP: "Doação de área / DUP (Declaração de Utilidade Pública)",
  areaTombada: "Área tombada ou envoltória",
  anuenciaMetro: "Anuência do Metrô",
};

const REGISTRO_TIPOS = ["CREA", "CAU"];
const VINCULO_TIPOS = ["Projeto", "Execução", "Projeto e Execução"];
function novoResponsavelTecnico() {
  return { id: Date.now() + Math.random(), vinculo: "Projeto", nome: "", cpf: "", registroTipo: "CREA", registroNumero: "", ccm: "", art: "", carteiraAnexada: false, artAnexada: false };
}

/* ============================================================
   FINANCEIRO — faturamento por processo/serviço
   ============================================================ */
const FATURAMENTO_CONFIG = {
  nao_faturado: { label: "Não faturado", fg: COLORS.steel, bg: "rgba(255,255,255,0.06)" },
  faturado: { label: "Faturado", fg: COLORS.orange, bg: COLORS.orangeDim },
  pago: { label: "Pago", fg: COLORS.green, bg: COLORS.greenDim },
};
const FATURAMENTO_KEYS = Object.keys(FATURAMENTO_CONFIG);
function isPassivo(processo) {
  return (processo.statusFaturamento === "faturado" || processo.statusFaturamento === "pago") && !STATUS_CONFIG[processo.statusAtual].final;
}
function processoBloqueado(processo, todosProcessos) {
  if (!processo.dependeDeId) return null;
  const dep = todosProcessos.find((p) => p.id === processo.dependeDeId);
  if (!dep) return null;
  return STATUS_CONFIG[dep.statusAtual].final ? null : dep;
}
function mesLabel(iso) {
  if (!iso) return null;
  const [y, m] = iso.split("-");
  const nomes = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
  return `${nomes[parseInt(m, 10) - 1]}/${y.slice(2)}`;
}
const MESES_ABREV = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
const MESES_NOMES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

/* ============================================================
   DADOS DE EXEMPLO (mock)
   ============================================================ */
const hoje = new Date();
const d = (offsetDays) => { const dt = new Date(hoje); dt.setDate(dt.getDate() + offsetDays); return dt.toISOString().slice(0, 10); };

function baseProcesso(p) {
  return {
    id: p.id, cliente: p.cliente, unidade: p.unidade, cidade: p.cidade, uf: p.uf,
    assunto: p.assunto, tipo: p.tipo || "Processo", numero: p.numero || "-",
    statusAtual: p.statusAtual || "aguardando", prestador: p.prestador || "-", tecnico: p.tecnico || "-",
    dataInicio: p.dataInicio || null, dataPrevisaoAnaliseChecklist: p.dataPrevisaoAnaliseChecklist || null,
    dataPrevistaProtocolo: p.dataPrevistaProtocolo || null,
    dataProtocolo: p.dataProtocolo || null, dataPrevisaoOrgao: p.dataPrevisaoOrgao || null,
    dataExigenciaRecebida: p.dataExigenciaRecebida || null, dataExigenciaPrazoLimite: p.dataExigenciaPrazoLimite || null,
    dataAtendimentoTecnico: p.dataAtendimentoTecnico || null, dataAtendimentoExigencia: p.dataAtendimentoExigencia || null,
    dataConclusao: p.dataConclusao || null,
    dataPrevistaVistoria: p.dataPrevistaVistoria || null, dataPrevisaoEntrega: p.dataPrevisaoEntrega || null,
    cobrancas: p.cobrancas || [], pendenciaCliente: p.pendenciaCliente || { ativa: false, descricao: "" },
    ultimaAtualizacao: p.ultimaAtualizacao || d(0),
    site: p.site || "-", login: p.login || "-", senha: p.senha || "-",
    checklist: p.checklist || instantiateChecklist(pickTemplateId(p.assunto)),
    documentos: p.documentos || instantiateDocumentos(pickTemplateId(p.assunto)),
    responsaveisTecnicos: p.responsaveisTecnicos || [],
    parametrosUrbanisticos: p.parametrosUrbanisticos || emptyParametrosUrbanisticos(),
    enquadramentos: p.enquadramentos || emptyEnquadramentos(),
    dependeDeId: p.dependeDeId || null,
    numeroContrato: p.numeroContrato || "-",
    valorContrato: p.valorContrato || 0,
    statusFaturamento: p.statusFaturamento || "nao_faturado",
    dataFaturamento: p.dataFaturamento || null,
    parcelasContrato: p.parcelasContrato || [],
    atualizacoes: p.atualizacoes || [],
  };
}

const MOCK_PROCESSOS = [];

/* ============================================================
   CONVERSÃO BANCO ↔ APP — o banco usa snake_case (padrão do
   Postgres), o app usa camelCase. Essas funções traduzem nos
   dois sentidos, em um único lugar.
   ============================================================ */
function rowToProcesso(r) {
  return baseProcesso({
    id: r.id, cliente: r.cliente, unidade: r.unidade, cidade: r.cidade, uf: r.uf,
    assunto: r.assunto, tipo: r.tipo, numero: r.numero, tecnico: r.tecnico,
    statusAtual: r.status_atual, prestador: r.prestador,
    numeroContrato: r.numero_contrato, valorContrato: r.valor_contrato,
    statusFaturamento: r.status_faturamento, dataFaturamento: r.data_faturamento,
    dataInicio: r.data_inicio, dataPrevisaoAnaliseChecklist: r.data_previsao_analise_checklist,
    dataPrevistaProtocolo: r.data_prevista_protocolo, dataProtocolo: r.data_protocolo,
    dataPrevisaoOrgao: r.data_previsao_orgao,
    dataExigenciaRecebida: r.data_exigencia_recebida, dataExigenciaPrazoLimite: r.data_exigencia_prazo_limite,
    dataAtendimentoTecnico: r.data_atendimento_tecnico, dataAtendimentoExigencia: r.data_atendimento_exigencia,
    dataConclusao: r.data_conclusao, dataPrevistaVistoria: r.data_prevista_vistoria, dataPrevisaoEntrega: r.data_previsao_entrega,
    ultimaAtualizacao: r.ultima_atualizacao,
    site: r.site, login: r.login, senha: r.senha,
    dependeDeId: r.depende_de_id,
    checklist: (r.checklist && r.checklist.secoes) ? r.checklist : undefined,
    documentos: (r.documentos && r.documentos.categorias) ? r.documentos : undefined,
    responsaveisTecnicos: r.responsaveis_tecnicos || [],
    parametrosUrbanisticos: (r.parametros_urbanisticos && Object.keys(r.parametros_urbanisticos).length) ? r.parametros_urbanisticos : undefined,
    enquadramentos: (r.enquadramentos && Object.keys(r.enquadramentos).length) ? r.enquadramentos : undefined,
    cobrancas: r.cobrancas || [],
    pendenciaCliente: r.pendencia_cliente || { ativa: false, descricao: "" },
    parcelasContrato: r.parcelas_contrato || [],
    atualizacoes: r.atualizacoes || [],
  });
}
function processoToRow(p) {
  return {
    id: p.id, cliente: p.cliente, unidade: p.unidade, cidade: p.cidade, uf: p.uf,
    assunto: p.assunto, tipo: p.tipo, numero: p.numero, tecnico: p.tecnico,
    status_atual: p.statusAtual, prestador: p.prestador,
    numero_contrato: p.numeroContrato, valor_contrato: p.valorContrato,
    status_faturamento: p.statusFaturamento, data_faturamento: p.dataFaturamento || null,
    data_inicio: p.dataInicio || null, data_previsao_analise_checklist: p.dataPrevisaoAnaliseChecklist || null,
    data_prevista_protocolo: p.dataPrevistaProtocolo || null, data_protocolo: p.dataProtocolo || null,
    data_previsao_orgao: p.dataPrevisaoOrgao || null,
    data_exigencia_recebida: p.dataExigenciaRecebida || null, data_exigencia_prazo_limite: p.dataExigenciaPrazoLimite || null,
    data_atendimento_tecnico: p.dataAtendimentoTecnico || null, data_atendimento_exigencia: p.dataAtendimentoExigencia || null,
    data_conclusao: p.dataConclusao || null, data_prevista_vistoria: p.dataPrevistaVistoria || null, data_previsao_entrega: p.dataPrevisaoEntrega || null,
    ultima_atualizacao: p.ultimaAtualizacao || null,
    site: p.site, login: p.login, senha: p.senha,
    depende_de_id: p.dependeDeId,
    checklist: p.checklist, documentos: p.documentos,
    responsaveis_tecnicos: p.responsaveisTecnicos,
    parametros_urbanisticos: p.parametrosUrbanisticos,
    enquadramentos: p.enquadramentos,
    cobrancas: p.cobrancas,
    pendencia_cliente: p.pendenciaCliente,
    parcelas_contrato: p.parcelasContrato,
    atualizacoes: p.atualizacoes,
  };
}
function rowToContrato(r) {
  return {
    id: r.id, proposta: r.proposta, cliente: r.cliente, unidade: r.unidade, codigoLoja: r.codigo_loja,
    servico: r.servico, tarefa: r.tarefa, tecnico: r.tecnico, coordenador: r.coordenador, tipo: r.tipo,
    honorarios: r.honorarios, valorFaturamento: r.valor_faturamento, porcentagem: r.porcentagem,
    dataSLA: r.data_sla, dataFaturamento: r.data_faturamento, dataSLAServico: r.data_sla_servico,
    statusContrato: r.status_contrato, statusServico: r.status_servico, statusParcela: r.status_parcela,
    observacao: r.observacao,
  };
}
function contratoToRow(c) {
  return {
    id: c.id, proposta: c.proposta, cliente: c.cliente, unidade: c.unidade, codigo_loja: c.codigoLoja,
    servico: c.servico, tarefa: c.tarefa, tecnico: c.tecnico, coordenador: c.coordenador, tipo: c.tipo,
    honorarios: c.honorarios, valor_faturamento: c.valorFaturamento, porcentagem: c.porcentagem,
    data_sla: c.dataSLA || null, data_faturamento: c.dataFaturamento || null, data_sla_servico: c.dataSLAServico || null,
    status_contrato: c.statusContrato, status_servico: c.statusServico, status_parcela: c.statusParcela,
    observacao: c.observacao,
  };
}
function rowToAgendaItem(r) {
  return { id: r.id, data: r.data, titulo: r.titulo, tipo: r.tipo, tecnico: r.tecnico, descricao: r.descricao };
}

/* Chama o backend seguro (Edge Function) que cria/lista/remove acessos.
   Usa o token da sessão atual — nunca a chave secreta. */
async function callAdminUsers(action, payload) {
  const { data: sessionData } = await supabase.auth.getSession();
  const token = sessionData?.session?.access_token;
  const resp = await fetch(`${SUPABASE_URL}/functions/v1/admin-users`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
    body: JSON.stringify({ action, ...payload }),
  });
  return resp.json();
}


/* ============================================================
   HELPERS
   ============================================================ */
function diasRestantes(proc) {
  if (!proc.dataPrevisaoOrgao) return null;
  return Math.round((new Date(proc.dataPrevisaoOrgao) - hoje) / 86400000);
}
function diasSemAtualizacao(proc) {
  if (!proc.ultimaAtualizacao) return null;
  return Math.round((hoje - new Date(proc.ultimaAtualizacao)) / 86400000);
}
function prazoInfo(dias) {
  if (dias === null) return { label: "—", fg: COLORS.steel, bg: "rgba(255,255,255,0.05)" };
  if (dias < 0) return { label: `Vencido (${Math.abs(dias)}d)`, fg: "#ffb3ac", bg: COLORS.overdueDim };
  if (dias <= 5) return { label: `${dias}d restantes`, fg: COLORS.red, bg: COLORS.redDim };
  if (dias <= 15) return { label: `${dias}d restantes`, fg: COLORS.orange, bg: COLORS.orangeDim };
  if (dias <= 30) return { label: `${dias}d restantes`, fg: COLORS.yellow, bg: COLORS.yellowDim };
  return { label: `${dias}d restantes`, fg: COLORS.green, bg: COLORS.greenDim };
}
function fmtDate(iso) {
  if (!iso) return "—";
  const [y, m, dd] = iso.split("-");
  return `${dd}/${m}/${y}`;
}
let OCULTAR_VALORES = false; // controlado pelo papel do usuário logado (Janayna não vê valores)
function fmtBRL(v) {
  if (OCULTAR_VALORES) return "••••••";
  const n = Number(v) || 0;
  return `R$ ${n.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}
function fmtDateTime(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  return `${fmtDate(iso.slice(0, 10))} às ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}
function checklistProgress(checklist) {
  let total = 0, ok = 0;
  checklist.secoes.forEach((s) => s.itens.forEach((it) => {
    if (it.estado === "na") return;
    total += 1;
    if (it.estado === "conforme") ok += 1;
  }));
  return total === 0 ? 100 : Math.round((ok / total) * 100);
}
function csvEscape(v) {
  const s = (v === null || v === undefined) ? "" : String(v);
  return `"${s.replace(/"/g, '""')}"`;
}
function downloadCSV(filename, rows) {
  const content = rows.map((r) => r.map(csvEscape).join(";")).join("\n");
  const blob = new Blob(["\uFEFF" + content], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
function fmtBytes(n) {
  if (!n && n !== 0) return "";
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

/* Monta a lista de pendências que bloqueiam um protocolo de qualidade */
function pendenciasProtocolo(processo) {
  const pendencias = [];
  processo.documentos.categorias.forEach((c) => c.itens.forEach((it) => {
    if (it.aplicavel && it.obrigatorio && it.estado === "pendente") pendencias.push(`Documento pendente: ${it.texto} (${c.titulo})`);
  }));
  processo.checklist.secoes.forEach((s) => s.itens.forEach((it) => {
    if (it.estado === "nao_conforme") pendencias.push(`Checklist técnico não conforme: ${it.texto} (${s.titulo})`);
    if (it.estado === "pendente") pendencias.push(`Checklist técnico pendente: ${it.texto} (${s.titulo})`);
  }));
  if (processo.responsaveisTecnicos.length === 0) pendencias.push("Nenhum responsável técnico cadastrado.");
  processo.responsaveisTecnicos.forEach((r) => {
    if (!r.nome || !r.registroNumero || !r.art) pendencias.push(`Dados incompletos do responsável técnico "${r.nome || "sem nome"}".`);
    if (!r.carteiraAnexada) pendencias.push(`Carteira profissional não anexada — ${r.nome || "responsável sem nome"}.`);
    if (!r.artAnexada) pendencias.push(`ART/RRT não anexada — ${r.nome || "responsável sem nome"}.`);
  });
  Object.entries(processo.enquadramentos).forEach(([k, v]) => {
    if (v.aplica && !v.obs) pendencias.push(`Enquadramento "${ENQUADRAMENTOS_LABELS[k]}" aplicável sem observação/anexo registrado.`);
  });
  if (processo.pendenciaCliente.ativa) pendencias.push(`Pendência do cliente em aberto: ${processo.pendenciaCliente.descricao || "sem descrição"}.`);
  return pendencias;
}

function gerarRelatorioHTML(processo) {
  const docPct = documentosProgress(processo.documentos);
  const chkPct = checklistProgress(processo.checklist);
  const pendencias = pendenciasProtocolo(processo);
  const pronto = pendencias.length === 0;
  const linhaDoc = (it) => `<tr><td>${it.texto}</td><td>${it.estado === "anexado" ? "Anexado" : it.estado === "nao_aplica" ? "Não se aplica" : "Pendente"}</td><td>${it.arquivo ? it.arquivo.nome : ""}</td></tr>`;
  const linhaChk = (it) => `<tr><td>${it.texto}</td><td>${it.estado === "conforme" ? "Conforme" : it.estado === "nao_conforme" ? "Não conforme" : it.estado === "na" ? "Não se aplica" : "Pendente"}</td><td>${it.obs || ""}</td></tr>`;
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Relatório — ${processo.assunto}</title>
  <style>
    body{font-family:Arial,Helvetica,sans-serif;color:#111;padding:32px;max-width:900px;margin:0 auto;}
    h1{font-size:20px;border-bottom:3px solid #e1483d;padding-bottom:10px;}
    h2{font-size:14px;text-transform:uppercase;letter-spacing:.04em;color:#333;margin-top:26px;border-left:4px solid #16283d;padding-left:8px;}
    table{width:100%;border-collapse:collapse;font-size:12px;margin-top:8px;}
    td,th{border:1px solid #ccc;padding:6px 8px;text-align:left;}
    th{background:#f0f0f0;}
    .badge{display:inline-block;padding:3px 10px;border-radius:3px;font-size:11px;font-weight:bold;margin-right:6px;}
    .ok{background:#dff5e8;color:#117a45;} .warn{background:#fdece9;color:#a3261b;}
    .pend{margin:4px 0;font-size:12.5px;} .meta{font-size:12px;color:#555;margin-top:4px;}
  </style></head><body>
  <h1>Relatório de conformidade — ${processo.assunto}</h1>
  <div class="meta">${processo.cliente} — ${processo.unidade} · ${processo.cidade}/${processo.uf} · Processo nº ${processo.numero}</div>
  <div class="meta">Gerado em ${fmtDate(new Date().toISOString().slice(0,10))} · Primers Consultoria e Legalização Imobiliária</div>
  <p style="margin-top:16px;">
    <span class="badge ${pronto ? "ok" : "warn"}">${pronto ? "Pronto para protocolo" : `${pendencias.length} pendência(s) para protocolo`}</span>
    <span class="badge ${docPct === 100 ? "ok" : "warn"}">Documentos ${docPct}%</span>
    <span class="badge ${chkPct === 100 ? "ok" : "warn"}">Checklist técnico ${chkPct}%</span>
  </p>
  ${pendencias.length > 0 ? `<h2>Pendências identificadas</h2>${pendencias.map((p) => `<div class="pend">• ${p}</div>`).join("")}` : ""}
  <h2>Documentos exigidos</h2>
  ${processo.documentos.categorias.map((c) => `<b>${c.titulo}</b><table><tr><th>Item</th><th>Status</th><th>Arquivo</th></tr>${c.itens.filter((i) => i.aplicavel).map(linhaDoc).join("")}</table>`).join("")}
  <h2>Checklist técnico</h2>
  ${processo.checklist.secoes.map((s) => `<b>${s.titulo}</b><table><tr><th>Item</th><th>Status</th><th>Observação</th></tr>${s.itens.map(linhaChk).join("")}</table>`).join("")}
  <h2>Responsáveis técnicos</h2>
  <table><tr><th>Vínculo</th><th>Nome</th><th>CPF</th><th>Registro</th><th>CCM</th><th>ART/RRT</th></tr>
  ${processo.responsaveisTecnicos.map((r) => `<tr><td>${r.vinculo}</td><td>${r.nome}</td><td>${r.cpf}</td><td>${r.registroTipo} ${r.registroNumero}</td><td>${r.ccm}</td><td>${r.art}</td></tr>`).join("") || `<tr><td colspan="6">Nenhum responsável técnico cadastrado.</td></tr>`}
  </table>
  <h2>Parâmetros urbanísticos</h2>
  <table>${PARAM_URB_FIELDS.map(([k, label]) => `<tr><td>${label}</td><td>${processo.parametrosUrbanisticos[k] || "—"}</td></tr>`).join("")}</table>
  <h2>Enquadramentos especiais</h2>
  <table><tr><th>Enquadramento</th><th>Aplica?</th><th>Observação</th></tr>
  ${Object.entries(processo.enquadramentos).map(([k, v]) => `<tr><td>${ENQUADRAMENTOS_LABELS[k]}</td><td>${v.aplica ? "Sim" : "Não"}</td><td>${v.obs || ""}</td></tr>`).join("")}
  </table>
  </body></html>`;
}
function imprimirRelatorio(processo) {
  const win = window.open("", "_blank");
  if (!win) return;
  win.document.write(gerarRelatorioHTML(processo));
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 300);
}

/* ============================================================
   ATOMS
   ============================================================ */
function Pill({ children, fg, bg, stamp }) {
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 5,
      padding: stamp ? "3px 10px" : "3px 9px", borderRadius: stamp ? 3 : 999,
      fontSize: 11, fontWeight: 700, letterSpacing: stamp ? "0.06em" : "0.02em",
      textTransform: stamp ? "uppercase" : "none", color: fg, background: bg,
      border: stamp ? `1px solid ${fg}55` : "1px solid transparent",
      fontFamily: stamp ? "'Oswald', sans-serif" : "'Inter', sans-serif", whiteSpace: "nowrap",
    }}>{children}</span>
  );
}

function KpiCard({ icon: Icon, label, value, accent, sub }) {
  return (
    <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: "14px 16px", flex: 1, minWidth: 140, position: "relative", overflow: "hidden" }}>
      <div style={{ position: "absolute", right: -10, top: -10, opacity: 0.08 }}><Icon size={64} color={accent} /></div>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
        <div style={{ width: 24, height: 24, borderRadius: 6, background: `${accent}22`, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Icon size={13} color={accent} />
        </div>
        <span style={{ fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600 }}>{label}</span>
      </div>
      <div style={{ fontFamily: "'Oswald', sans-serif", fontSize: 26, fontWeight: 600, color: COLORS.ice, lineHeight: 1 }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: COLORS.steel, marginTop: 5 }}>{sub}</div>}
    </div>
  );
}

function Select({ value, onChange, options, placeholder }) {
  return (
    <select value={value} onChange={(e) => onChange(e.target.value)} style={{
      background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 8,
      padding: "8px 10px", color: COLORS.steelLight, fontSize: 12.5, fontFamily: "'Inter', sans-serif",
    }}>
      {placeholder && <option value="Todos">{placeholder}</option>}
      {options.map((o) => <option key={o} value={o}>{o}</option>)}
    </select>
  );
}

/* ============================================================
   FILTER BAR (global — usada no Dashboard e em Processos)
   ============================================================ */
function FilterBar({ processos, filtros, setFiltros }) {
  const clientes = useMemo(() => ["Todos", ...Array.from(new Set(processos.map((p) => p.cliente))).sort()], [processos]);
  const unidades = useMemo(() => {
    const base = filtros.cliente === "Todos" ? processos : processos.filter((p) => p.cliente === filtros.cliente);
    return ["Todos", ...Array.from(new Set(base.map((p) => p.unidade))).sort()];
  }, [processos, filtros.cliente]);
  const assuntos = useMemo(() => ["Todos", ...Array.from(new Set(processos.map((p) => p.assunto))).sort()], [processos]);

  const set = (k) => (v) => setFiltros((f) => ({ ...f, [k]: v, ...(k === "cliente" ? { unidade: "Todos" } : {}) }));

  return (
    <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 18 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, color: COLORS.steel, fontSize: 11.5, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.04em" }}>
        <Filter size={13} /> Filtros
      </div>
      <Select value={filtros.cliente} onChange={set("cliente")} options={clientes.filter((c) => c !== "Todos")} placeholder="Todos os clientes" />
      <Select value={filtros.unidade} onChange={set("unidade")} options={unidades.filter((u) => u !== "Todos")} placeholder="Todas as unidades" />
      <Select value={filtros.assunto} onChange={set("assunto")} options={assuntos.filter((a) => a !== "Todos")} placeholder="Todos os tipos de serviço" />
      <Select value={filtros.responsavel} onChange={set("responsavel")} options={RESPONSAVEIS} placeholder="Toda responsabilidade" />
      {(filtros.cliente !== "Todos" || filtros.unidade !== "Todos" || filtros.assunto !== "Todos" || filtros.responsavel !== "Todos") && (
        <button onClick={() => setFiltros({ cliente: "Todos", unidade: "Todos", assunto: "Todos", responsavel: "Todos" })}
          style={{ background: "transparent", border: "none", color: COLORS.red, fontSize: 12, cursor: "pointer", textDecoration: "underline" }}>
          Limpar filtros
        </button>
      )}
    </div>
  );
}

function applyFiltros(processos, filtros) {
  return processos.filter((p) => {
    if (filtros.cliente !== "Todos" && p.cliente !== filtros.cliente) return false;
    if (filtros.unidade !== "Todos" && p.unidade !== filtros.unidade) return false;
    if (filtros.assunto !== "Todos" && p.assunto !== filtros.assunto) return false;
    if (filtros.responsavel !== "Todos" && STATUS_CONFIG[p.statusAtual].responsavel !== filtros.responsavel) return false;
    return true;
  });
}

/* ============================================================
   NEW PROCESS MODAL
   ============================================================ */
function NewProcessModal({ onClose, onSave, processos, isAdmin }) {
  const [form, setForm] = useState({
    cliente: "", unidade: "", cidade: "", uf: "", assunto: "", tipo: "Processo",
    numero: "", statusAtual: "aguardando", dataProtocolo: "", dataPrevisaoOrgao: "",
    dataAtendimentoExigencia: "", pendenciaClienteDescricao: "",
    site: "", login: "", senha: "", prestador: "Interno",
    numeroContrato: "", valorContrato: "", statusFaturamento: "nao_faturado", dependeDeId: "",
  });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const field = (label, key, placeholder, type = "text") => (
    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      <label style={{ fontSize: 11, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 600 }}>{label}</label>
      <input type={type} value={form[key]} onChange={set(key)} placeholder={placeholder}
        style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "8px 10px", color: COLORS.ice, fontSize: 13, fontFamily: "'Inter', sans-serif", outline: "none" }} />
    </div>
  );
  const selectField = (label, key, options, getLabel) => (
    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      <label style={{ fontSize: 11, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 600 }}>{label}</label>
      <select value={form[key]} onChange={set(key)} style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "8px 10px", color: COLORS.ice, fontSize: 13, fontFamily: "'Inter', sans-serif", outline: "none" }}>
        {options.map((o) => <option key={o} value={o}>{getLabel ? getLabel(o) : o}</option>)}
      </select>
    </div>
  );

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(5,10,16,0.7)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }} onClick={onClose}>
      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.borderStrong}`, borderRadius: 12, width: "100%", maxWidth: 680, maxHeight: "88vh", overflowY: "auto", padding: 24 }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
          <h2 style={{ fontFamily: "'Oswald', sans-serif", fontSize: 19, fontWeight: 600, color: COLORS.ice, textTransform: "uppercase", letterSpacing: "0.03em" }}>Novo processo</h2>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.steel }}><X size={20} /></button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          {field("Cliente", "cliente", "Ex: Atacado Boreal")}
          {field("Unidade", "unidade", "Ex: Loja Centro")}
          {field("Cidade", "cidade", "Ex: São Paulo")}
          {field("UF", "uf", "Ex: SP")}
          <div style={{ gridColumn: "1 / -1" }}>{field("Assunto / Serviço", "assunto", "Ex: Aprovação de Projeto - Prefeitura (Obra Nova)")}</div>
          {selectField("Tipo de serviço", "tipo", CONTRATO_TIPO_OPTIONS)}
          {field("Nº do processo", "numero", "Ex: 1101.2025/0001")}
          <div style={{ gridColumn: "1 / -1" }}>{selectField("Status atual (responsabilidade)", "statusAtual", STATUS_KEYS, (k) => `${STATUS_CONFIG[k].label} — ${STATUS_CONFIG[k].responsavel}`)}</div>
          {field("Data de protocolo", "dataProtocolo", "", "date")}
          {field("Previsão de análise do órgão", "dataPrevisaoOrgao", "", "date")}
          {field("Data de atendimento de exigência", "dataAtendimentoExigencia", "", "date")}
          {field("Pendência do cliente (se houver)", "pendenciaClienteDescricao", "Ex: aguardando envio de documento X")}
          {field("Site do órgão", "site", "portal.orgao.gov.br")}
          {field("Prestador responsável", "prestador", "Interno / nome")}
          {field("Login", "login", "usuário do portal")}
          {field("Senha", "senha", "senha do portal", "password")}
          {field("Nº do contrato", "numeroContrato", "Ex: CT-2026-0001")}
          {isAdmin && field("Valor do serviço (R$)", "valorContrato", "Ex: 4500", "number")}
          {selectField("Status de faturamento", "statusFaturamento", FATURAMENTO_KEYS, (k) => FATURAMENTO_CONFIG[k].label)}
          <div style={{ gridColumn: "1 / -1" }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
              <label style={{ fontSize: 11, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 600 }}>Depende da conclusão de outro processo? (opcional)</label>
              <select value={form.dependeDeId} onChange={(e) => setForm((f) => ({ ...f, dependeDeId: e.target.value }))}
                style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "8px 10px", color: COLORS.ice, fontSize: 13, fontFamily: "'Inter', sans-serif", outline: "none" }}>
                <option value="">Nenhuma dependência</option>
                {processos.filter((p) => !STATUS_CONFIG[p.statusAtual].final).map((p) => (
                  <option key={p.id} value={p.id}>{p.cliente} — {p.assunto}</option>
                ))}
              </select>
            </div>
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 22 }}>
          <button onClick={onClose} style={{ background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "9px 16px", fontSize: 13, cursor: "pointer" }}>Cancelar</button>
          <button onClick={() => {
            if (!form.cliente || !form.assunto) return;
            const pendenciaCliente = form.pendenciaClienteDescricao ? { ativa: true, descricao: form.pendenciaClienteDescricao } : { ativa: false, descricao: "" };
            onSave(baseProcesso({
              ...form, id: Date.now(), pendenciaCliente, ultimaAtualizacao: new Date().toISOString().slice(0, 10),
              valorContrato: parseFloat(form.valorContrato) || 0, dependeDeId: form.dependeDeId ? Number(form.dependeDeId) : null,
            }));
            onClose();
          }} style={{ background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
            Salvar processo
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   CHECKLIST TAB
   ============================================================ */
function ChecklistTab({ processo, onUpdate }) {
  const setItemState = (si, ii, estado) => {
    const checklist = JSON.parse(JSON.stringify(processo.checklist));
    checklist.secoes[si].itens[ii].estado = checklist.secoes[si].itens[ii].estado === estado ? "pendente" : estado;
    onUpdate({ ...processo, checklist });
  };
  const setItemObs = (si, ii, obs) => {
    const checklist = JSON.parse(JSON.stringify(processo.checklist));
    checklist.secoes[si].itens[ii].obs = obs;
    onUpdate({ ...processo, checklist });
  };
  const templateNome = (CHECKLIST_TEMPLATES[processo.checklist.templateId] || CHECKLIST_TEMPLATES.default).nome;
  const progresso = checklistProgress(processo.checklist);

  const StateBtn = ({ active, color, icon: Icon, onClick, title }) => (
    <button onClick={onClick} title={title} style={{
      width: 26, height: 26, borderRadius: 6, border: `1px solid ${active ? color : COLORS.border}`,
      background: active ? `${color}22` : "transparent", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer",
    }}>
      <Icon size={14} color={active ? color : COLORS.steel} />
    </button>
  );

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <div>
          <div style={{ fontSize: 11, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em" }}>Template</div>
          <div style={{ fontSize: 14, color: COLORS.ice, fontWeight: 600, fontFamily: "'Oswald', sans-serif" }}>{templateNome}</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontSize: 11, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em" }}>Conformidade</div>
          <div style={{ fontSize: 20, color: progresso === 100 ? COLORS.green : COLORS.ice, fontWeight: 700, fontFamily: "'Oswald', sans-serif" }}>{progresso}%</div>
        </div>
      </div>
      {processo.checklist.secoes.map((sec, si) => {
        const secProg = checklistProgress({ secoes: [sec] });
        return (
          <div key={si} style={{ marginBottom: 18 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, color: COLORS.steelLight, textTransform: "uppercase", letterSpacing: "0.04em" }}>{sec.titulo}</div>
              <div style={{ fontSize: 11, color: COLORS.steel }}>{secProg}%</div>
            </div>
            <div style={{ height: 4, background: COLORS.panelAlt, borderRadius: 2, marginBottom: 10, overflow: "hidden" }}>
              <div style={{ height: "100%", width: `${secProg}%`, background: secProg === 100 ? COLORS.green : COLORS.orange }} />
            </div>
            {sec.itens.map((it, ii) => (
              <div key={it.id} style={{ padding: "9px 0", borderBottom: `1px solid ${COLORS.border}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 12.5, color: COLORS.ice }}>{it.texto}</div>
                    {it.detalhe && <div style={{ fontSize: 11, color: COLORS.steel, marginTop: 3, lineHeight: 1.5 }}>{it.detalhe}</div>}
                  </div>
                  <div style={{ display: "flex", gap: 5, flexShrink: 0 }}>
                    <StateBtn active={it.estado === "conforme"} color={COLORS.green} icon={CheckCircle2} title="Conforme" onClick={() => setItemState(si, ii, "conforme")} />
                    <StateBtn active={it.estado === "nao_conforme"} color={COLORS.red} icon={XCircle} title="Não conforme" onClick={() => setItemState(si, ii, "nao_conforme")} />
                    <StateBtn active={it.estado === "na"} color={COLORS.steel} icon={MinusCircle} title="Não se aplica" onClick={() => setItemState(si, ii, "na")} />
                  </div>
                </div>
                {it.estado === "nao_conforme" && (
                  <input value={it.obs} onChange={(e) => setItemObs(si, ii, e.target.value)} placeholder="Observação sobre a não conformidade..."
                    style={{ marginTop: 6, width: "100%", background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 9px", color: COLORS.ice, fontSize: 11.5, outline: "none" }} />
                )}
              </div>
            ))}
          </div>
        );
      })}
    </div>
  );
}

/* ============================================================
   DOCUMENTOS TAB — checklist documental com anexo, responsáveis
   técnicos e enquadramentos especiais
   ============================================================ */
function DocumentosTab({ processo, onUpdate }) {
  const patchDoc = (ci, ii, fields) => {
    const documentos = JSON.parse(JSON.stringify(processo.documentos));
    Object.assign(documentos.categorias[ci].itens[ii], fields);
    onUpdate({ ...processo, documentos });
  };
  const anexar = (ci, ii, file) => {
    if (!file) return;
    patchDoc(ci, ii, { estado: "anexado", arquivo: { nome: file.name, tamanho: file.size } });
  };
  const progresso = documentosProgress(processo.documentos);

  const addResponsavel = () => onUpdate({ ...processo, responsaveisTecnicos: [...processo.responsaveisTecnicos, novoResponsavelTecnico()] });
  const patchResponsavel = (id, fields) => onUpdate({ ...processo, responsaveisTecnicos: processo.responsaveisTecnicos.map((r) => (r.id === id ? { ...r, ...fields } : r)) });
  const removeResponsavel = (id) => onUpdate({ ...processo, responsaveisTecnicos: processo.responsaveisTecnicos.filter((r) => r.id !== id) });

  const patchEnquadramento = (k, fields) => onUpdate({ ...processo, enquadramentos: { ...processo.enquadramentos, [k]: { ...processo.enquadramentos[k], ...fields } } });

  return (
    <div>
      <div style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 14px", fontSize: 11.5, color: COLORS.steelLight, marginBottom: 18, lineHeight: 1.5 }}>
        Nesta fase do MVP os anexos ficam registrados pelo nome do arquivo para fins de conferência — o armazenamento definitivo dos documentos entra na versão integrada ao sistema.
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <div style={{ fontSize: 14, color: COLORS.ice, fontWeight: 600, fontFamily: "'Oswald', sans-serif" }}>
          {(DOC_TEMPLATES[processo.documentos.templateId] || DOC_TEMPLATES.default).nome}
        </div>
        <div style={{ fontSize: 20, color: progresso === 100 ? COLORS.green : COLORS.ice, fontWeight: 700, fontFamily: "'Oswald', sans-serif" }}>{progresso}%</div>
      </div>

      {processo.documentos.categorias.map((cat, ci) => (
        <div key={ci} style={{ marginBottom: 18 }}>
          <div style={{ fontSize: 12.5, fontWeight: 700, color: COLORS.steelLight, textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 8 }}>{cat.titulo}</div>
          {cat.itens.map((it, ii) => (
            <div key={it.id} style={{ padding: "9px 0", borderBottom: `1px solid ${COLORS.border}` }}>
              {it.condicionalLabel && (
                <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6, cursor: "pointer" }}>
                  <input type="checkbox" checked={it.aplicavel} onChange={(e) => patchDoc(ci, ii, { aplicavel: e.target.checked, estado: e.target.checked ? "pendente" : "nao_aplica" })} />
                  <span style={{ fontSize: 11.5, color: COLORS.steel }}>{it.condicionalLabel}</span>
                </label>
              )}
              {it.aplicavel && (
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 12.5, color: COLORS.ice }}>{it.texto}{it.obrigatorio && <span style={{ color: COLORS.red }}> *</span>}</div>
                      {it.detalhe && <div style={{ fontSize: 11, color: COLORS.steel, marginTop: 3 }}>{it.detalhe}</div>}
                      {it.arquivo && <div style={{ fontSize: 11, color: COLORS.green, marginTop: 4 }}>{it.arquivo.nome} ({fmtBytes(it.arquivo.tamanho)})</div>}
                    </div>
                    <div style={{ display: "flex", gap: 6, flexShrink: 0, alignItems: "center" }}>
                      <label style={{ display: "flex", alignItems: "center", gap: 5, background: it.estado === "anexado" ? COLORS.greenDim : "transparent", border: `1px solid ${it.estado === "anexado" ? COLORS.green : COLORS.border}`, color: it.estado === "anexado" ? COLORS.green : COLORS.steelLight, borderRadius: 6, padding: "5px 10px", fontSize: 11.5, cursor: "pointer" }}>
                        <Upload size={12} /> {it.estado === "anexado" ? "Substituir" : "Anexar"}
                        <input type="file" style={{ display: "none" }} onChange={(e) => anexar(ci, ii, e.target.files[0])} />
                      </label>
                      <button onClick={() => patchDoc(ci, ii, { estado: it.estado === "nao_aplica" ? "pendente" : "nao_aplica" })}
                        title="Não se aplica" style={{ width: 26, height: 26, borderRadius: 6, border: `1px solid ${it.estado === "nao_aplica" ? COLORS.steel : COLORS.border}`, background: it.estado === "nao_aplica" ? COLORS.grayDim : "transparent", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" }}>
                        <MinusCircle size={13} color={it.estado === "nao_aplica" ? COLORS.steelLight : COLORS.steel} />
                      </button>
                    </div>
                  </div>
                  {it.estado === "anexado" && (
                    <div style={{ display: "flex", gap: 8, alignItems: "center", marginTop: 8, flexWrap: "wrap" }}>
                      <span style={{ fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.03em" }}>Validade:</span>
                      {["valido", "invalido"].map((v) => (
                        <button key={v} onClick={() => patchDoc(ci, ii, { validade: it.validade === v ? null : v })}
                          style={{
                            fontSize: 11, fontWeight: 700, borderRadius: 999, padding: "3px 10px", cursor: "pointer",
                            background: it.validade === v ? (v === "valido" ? COLORS.greenDim : COLORS.redDim) : "transparent",
                            color: it.validade === v ? (v === "valido" ? COLORS.green : COLORS.red) : COLORS.steel,
                            border: `1px solid ${it.validade === v ? (v === "valido" ? COLORS.green : COLORS.red) + "55" : COLORS.border}`,
                          }}>
                          {v === "valido" ? "Válido" : "Inválido"}
                        </button>
                      ))}
                      <input value={it.obs} onChange={(e) => patchDoc(ci, ii, { obs: e.target.value })} placeholder="Observação sobre o documento..."
                        style={{ flex: 1, minWidth: 160, background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "5px 9px", color: COLORS.ice, fontSize: 11.5 }} />
                    </div>
                  )}
                </div>
              )}
            </div>
          ))}
        </div>
      ))}

      <div style={{ marginTop: 22, display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <div style={{ fontSize: 12.5, fontWeight: 700, color: COLORS.steelLight, textTransform: "uppercase", letterSpacing: "0.04em" }}>Responsáveis técnicos</div>
        <button onClick={addResponsavel} style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 6, padding: "5px 10px", fontSize: 11.5, cursor: "pointer" }}>
          <Plus size={12} /> Adicionar responsável
        </button>
      </div>
      {processo.responsaveisTecnicos.length === 0 && <div style={{ fontSize: 12, color: COLORS.steel, padding: "8px 0 16px" }}>Nenhum responsável técnico cadastrado.</div>}
      {processo.responsaveisTecnicos.map((r) => (
        <div key={r.id} style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: 12, marginBottom: 10 }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 8 }}>
            <select value={r.vinculo} onChange={(e) => patchResponsavel(r.id, { vinculo: e.target.value })} style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 8px", color: COLORS.ice, fontSize: 12 }}>
              {VINCULO_TIPOS.map((v) => <option key={v} value={v}>{v}</option>)}
            </select>
            <input placeholder="Nome" value={r.nome} onChange={(e) => patchResponsavel(r.id, { nome: e.target.value })} style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 8px", color: COLORS.ice, fontSize: 12 }} />
            <input placeholder="CPF" value={r.cpf} onChange={(e) => patchResponsavel(r.id, { cpf: e.target.value })} style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 8px", color: COLORS.ice, fontSize: 12 }} />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "0.7fr 1fr 1fr 1fr", gap: 8, marginBottom: 8 }}>
            <select value={r.registroTipo} onChange={(e) => patchResponsavel(r.id, { registroTipo: e.target.value })} style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 8px", color: COLORS.ice, fontSize: 12 }}>
              {REGISTRO_TIPOS.map((v) => <option key={v} value={v}>{v}</option>)}
            </select>
            <input placeholder="Nº do registro" value={r.registroNumero} onChange={(e) => patchResponsavel(r.id, { registroNumero: e.target.value })} style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 8px", color: COLORS.ice, fontSize: 12 }} />
            <input placeholder="CCM" value={r.ccm} onChange={(e) => patchResponsavel(r.id, { ccm: e.target.value })} style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 8px", color: COLORS.ice, fontSize: 12 }} />
            <input placeholder="Nº da ART/RRT" value={r.art} onChange={(e) => patchResponsavel(r.id, { art: e.target.value })} style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 8px", color: COLORS.ice, fontSize: 12 }} />
          </div>
          <div style={{ display: "flex", gap: 14, alignItems: "center" }}>
            <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
              <input type="checkbox" checked={r.carteiraAnexada} onChange={(e) => patchResponsavel(r.id, { carteiraAnexada: e.target.checked })} />
              <span style={{ fontSize: 11.5, color: COLORS.steelLight }}>Carteira profissional anexada</span>
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
              <input type="checkbox" checked={r.artAnexada} onChange={(e) => patchResponsavel(r.id, { artAnexada: e.target.checked })} />
              <span style={{ fontSize: 11.5, color: COLORS.steelLight }}>ART/RRT anexada</span>
            </label>
            <button onClick={() => removeResponsavel(r.id)} style={{ marginLeft: "auto", background: "none", border: "none", color: COLORS.red, fontSize: 11.5, cursor: "pointer" }}>Remover</button>
          </div>
        </div>
      ))}

      <div style={{ marginTop: 22, fontSize: 12.5, fontWeight: 700, color: COLORS.steelLight, textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 10 }}>Enquadramentos especiais</div>
      {Object.entries(processo.enquadramentos).map(([k, v]) => (
        <div key={k} style={{ padding: "9px 0", borderBottom: `1px solid ${COLORS.border}` }}>
          <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", marginBottom: v.aplica ? 6 : 0 }}>
            <input type="checkbox" checked={v.aplica} onChange={(e) => patchEnquadramento(k, { aplica: e.target.checked })} />
            <span style={{ fontSize: 12.5, color: COLORS.ice }}>{ENQUADRAMENTOS_LABELS[k]}</span>
          </label>
          {v.aplica && (
            <input value={v.obs} onChange={(e) => patchEnquadramento(k, { obs: e.target.value })} placeholder="Observação / status do enquadramento..."
              style={{ width: "100%", background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "7px 9px", color: COLORS.ice, fontSize: 12 }} />
          )}
        </div>
      ))}
    </div>
  );
}

/* ============================================================
   PARÂMETROS URBANÍSTICOS TAB
   ============================================================ */
function ParametrosUrbanisticosTab({ processo, onUpdate }) {
  const set = (k) => (e) => onUpdate({ ...processo, parametrosUrbanisticos: { ...processo.parametrosUrbanisticos, [k]: e.target.value } });
  return (
    <div>
      <div style={{ fontSize: 11.5, color: COLORS.steel, marginBottom: 16, lineHeight: 1.5 }}>
        Quadro de uso e ocupação do solo do projeto — preencha conforme o memorial/quadro técnico elaborado para este serviço.
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        {PARAM_URB_FIELDS.map(([k, label]) => (
          <div key={k} style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.03em" }}>{label}</label>
            <input value={processo.parametrosUrbanisticos[k]} onChange={set(k)}
              style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "7px 9px", color: COLORS.ice, fontSize: 12.5 }} />
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================================================
   RELATÓRIO TAB — visão consolidada de qualidade do protocolo
   ============================================================ */
function RelatorioTab({ processo }) {
  const pendencias = pendenciasProtocolo(processo);
  const pronto = pendencias.length === 0;
  const docPct = documentosProgress(processo.documentos);
  const chkPct = checklistProgress(processo.checklist);

  const exportarCSV = () => {
    const rows = [["Categoria", "Item", "Status", "Detalhe"]];
    processo.documentos.categorias.forEach((c) => c.itens.filter((i) => i.aplicavel).forEach((it) => {
      rows.push(["Documento", `${c.titulo} — ${it.texto}`, it.estado === "anexado" ? "Anexado" : it.estado === "nao_aplica" ? "Não se aplica" : "Pendente", it.arquivo ? it.arquivo.nome : ""]);
    }));
    processo.checklist.secoes.forEach((s) => s.itens.forEach((it) => {
      rows.push(["Checklist técnico", `${s.titulo} — ${it.texto}`, it.estado === "conforme" ? "Conforme" : it.estado === "nao_conforme" ? "Não conforme" : it.estado === "na" ? "Não se aplica" : "Pendente", it.obs || ""]);
    }));
    downloadCSV(`relatorio_${processo.cliente}_${processo.numero}.csv`.replace(/[^\w.-]+/g, "_"), rows);
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 160, background: pronto ? COLORS.greenDim : COLORS.orangeDim, border: `1px solid ${pronto ? COLORS.green : COLORS.orange}55`, borderRadius: 10, padding: 14 }}>
          <div style={{ fontSize: 10.5, color: pronto ? COLORS.green : COLORS.orange, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700 }}>Status do protocolo</div>
          <div style={{ fontSize: 15, color: COLORS.ice, fontWeight: 700, fontFamily: "'Oswald', sans-serif", marginTop: 4 }}>{pronto ? "Pronto para protocolar" : `${pendencias.length} pendência(s)`}</div>
        </div>
        <div style={{ flex: 1, minWidth: 140, background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 14 }}>
          <div style={{ fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700 }}>Documentos</div>
          <div style={{ fontSize: 20, color: COLORS.ice, fontWeight: 700, fontFamily: "'Oswald', sans-serif", marginTop: 4 }}>{docPct}%</div>
        </div>
        <div style={{ flex: 1, minWidth: 140, background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 14 }}>
          <div style={{ fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700 }}>Checklist técnico</div>
          <div style={{ fontSize: 20, color: COLORS.ice, fontWeight: 700, fontFamily: "'Oswald', sans-serif", marginTop: 4 }}>{chkPct}%</div>
        </div>
      </div>

      <div style={{ display: "flex", gap: 10, marginBottom: 18 }}>
        <button onClick={() => imprimirRelatorio(processo)} style={{ display: "flex", alignItems: "center", gap: 7, background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.02em", textTransform: "uppercase" }}>
          Imprimir / salvar PDF
        </button>
        <button onClick={exportarCSV} style={{ display: "flex", alignItems: "center", gap: 7, background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "9px 14px", fontSize: 13, cursor: "pointer" }}>
          <Download size={14} /> Exportar CSV
        </button>
      </div>

      <div style={{ fontSize: 12.5, fontWeight: 700, color: COLORS.steelLight, textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 8 }}>Pendências identificadas</div>
      {pendencias.length === 0 ? (
        <div style={{ fontSize: 12.5, color: COLORS.green, padding: "8px 0" }}>Nenhuma pendência — documentos, checklist técnico e responsáveis técnicos completos.</div>
      ) : pendencias.map((p, i) => (
        <div key={i} style={{ fontSize: 12.5, color: COLORS.steelLight, padding: "6px 0", borderBottom: `1px solid ${COLORS.border}` }}>• {p}</div>
      ))}
    </div>
  );
}

/* ============================================================
   ATUALIZAÇÕES TAB (por processo)
   ============================================================ */
function AtualizacoesTab({ processo, onUpdate }) {
  const [form, setForm] = useState({ data: new Date().toISOString().slice(0, 10), tipo: ATUALIZACAO_TIPOS[0], descricao: "", responsavel: "Primers" });
  const add = () => {
    if (!form.descricao) return;
    const nova = { id: Date.now(), ...form };
    onUpdate({ ...processo, atualizacoes: [nova, ...processo.atualizacoes], ultimaAtualizacao: form.data });
    setForm((f) => ({ ...f, descricao: "" }));
  };
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
        <button onClick={() => imprimirStatusServico(processo)} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "7px 12px", fontSize: 12, cursor: "pointer" }}>
          <Download size={13} /> Exportar status de serviço
        </button>
      </div>
      <div style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: 14, marginBottom: 18 }}>
        <div style={{ fontSize: 11.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 600, marginBottom: 10 }}>Registrar atualização</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
          <input type="date" value={form.data} onChange={(e) => setForm((f) => ({ ...f, data: e.target.value }))}
            style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "7px 9px", color: COLORS.ice, fontSize: 12.5 }} />
          <select value={form.tipo} onChange={(e) => setForm((f) => ({ ...f, tipo: e.target.value }))}
            style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "7px 9px", color: COLORS.ice, fontSize: 12.5 }}>
            {ATUALIZACAO_TIPOS.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
          <select value={form.responsavel} onChange={(e) => setForm((f) => ({ ...f, responsavel: e.target.value }))}
            style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "7px 9px", color: COLORS.ice, fontSize: 12.5 }}>
            {["Primers", "Cliente", "Órgão"].map((r) => <option key={r} value={r}>{r}</option>)}
          </select>
        </div>
        <textarea value={form.descricao} onChange={(e) => setForm((f) => ({ ...f, descricao: e.target.value }))} placeholder="Descreva o que aconteceu..." rows={2}
          style={{ width: "100%", background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "8px 10px", color: COLORS.ice, fontSize: 12.5, resize: "vertical", fontFamily: "'Inter', sans-serif", marginBottom: 10 }} />
        <button onClick={add} style={{ display: "flex", alignItems: "center", gap: 6, background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "8px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.02em", textTransform: "uppercase" }}>
          <MessageSquarePlus size={14} /> Adicionar
        </button>
      </div>

      {processo.atualizacoes.length === 0 && <div style={{ fontSize: 12.5, color: COLORS.steel, textAlign: "center", padding: 20 }}>Nenhuma atualização registrada ainda.</div>}
      {processo.atualizacoes.map((a) => (
        <div key={a.id} style={{ display: "flex", gap: 12, padding: "10px 0", borderBottom: `1px solid ${COLORS.border}` }}>
          <div style={{ width: 74, flexShrink: 0, fontSize: 11.5, color: COLORS.steel, fontFamily: "monospace" }}>{fmtDate(a.data)}</div>
          <div style={{ flex: 1 }}>
            <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 3, flexWrap: "wrap" }}>
              <Pill fg={COLORS.steelLight} bg="rgba(255,255,255,0.06)">{a.tipo}</Pill>
              <span style={{ fontSize: 10.5, color: COLORS.steel }}>resp.: {a.responsavel}</span>
            </div>
            <div style={{ fontSize: 12.5, color: COLORS.ice, lineHeight: 1.5 }}>{a.descricao}</div>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ============================================================
   PROCESS DETAIL MODAL (Visão geral / Checklist / Atualizações)
   ============================================================ */
/* ============================================================
   PRÓXIMA AÇÃO GUIADA — cada status sabe qual é o próximo passo,
   quais datas pedir, e para qual status avançar. Adapta os
   rótulos conforme o serviço é Processo ou Serviço Técnico.
   ============================================================ */
function hojeISOStr() { return new Date().toISOString().slice(0, 10); }

function nextActionConfig(processo) {
  const tecnico = processo.tipo === "Serviço Técnico";
  switch (processo.statusAtual) {
    case "aguardando":
      return {
        titulo: "Iniciar serviço", textoBotao: "Iniciar Serviço",
        campos: [
          { key: "dataInicio", label: "Data de início", type: "date", default: hojeISOStr() },
          { key: "dataPrevisaoAnaliseChecklist", label: "Previsão de conclusão da análise documental / checklist", type: "date" },
        ],
        aplicar: (v) => ({ statusAtual: "iniciado", ...v }),
        resumo: () => "Serviço iniciado.",
      };
    case "iniciado":
      return {
        titulo: "Concluir análise e checklist", textoBotao: "Concluir análise/checklist",
        campos: [
          { key: tecnico ? "dataPrevistaVistoria" : "dataPrevistaProtocolo", label: tecnico ? "Data prevista de vistoria (se aplicável)" : "Data prevista de protocolo", type: "date" },
        ],
        aplicar: (v) => ({ statusAtual: "em_montagem", ...v }),
        resumo: () => "Análise documental e checklist concluídos — processo em montagem.",
      };
    case "em_montagem":
      return tecnico ? {
        titulo: "Registrar início da execução", textoBotao: "Registrar execução em andamento",
        campos: [
          { key: "dataPrevisaoEntrega", label: "Previsão de conclusão / entrega", type: "date" },
        ],
        aplicar: (v) => ({ statusAtual: "protocolado", ...v }),
        resumo: () => "Serviço técnico em execução.",
      } : {
        titulo: "Registrar protocolo", textoBotao: "Registrar protocolo",
        campos: [
          { key: "dataProtocolo", label: "Data de protocolo", type: "date", default: hojeISOStr() },
          { key: "dataPrevisaoOrgao", label: "Previsão de análise do órgão", type: "date" },
        ],
        aplicar: (v) => ({ statusAtual: "protocolado", ...v }),
        resumo: (v) => `Processo protocolado${v.dataProtocolo ? ` em ${fmtDate(v.dataProtocolo)}` : ""}.`,
      };
    case "protocolado":
    case "aguardando_orgao":
      return {
        titulo: tecnico ? "Registrar pendência ou ajuste" : "Registrar Comunique-se / Exigência",
        textoBotao: tecnico ? "Registrar pendência" : "Registrar exigência recebida",
        campos: [
          { key: "dataExigenciaRecebida", label: tecnico ? "Data da pendência" : "Data de recebimento da exigência", type: "date", default: hojeISOStr() },
          { key: "dataExigenciaPrazoLimite", label: "Data limite para atendimento", type: "date" },
        ],
        aplicar: (v) => ({ statusAtual: "exigencia_primers", ...v }),
        resumo: () => tecnico ? "Pendência registrada." : "Comunique-se / exigência recebida.",
      };
    case "exigencia_primers":
    case "exigencia_cliente":
      return {
        titulo: "Registrar atendimento", textoBotao: "Marcar exigência atendida",
        campos: [
          { key: "dataAtendimentoTecnico", label: "Data de atendimento técnico (esclarecimentos)", type: "date" },
          { key: "dataAtendimentoExigencia", label: "Data em que foi atendida", type: "date", default: hojeISOStr() },
        ],
        aplicar: (v) => ({ statusAtual: "protocolado", ...v }),
        resumo: () => "Exigência atendida — processo volta a Protocolado / aguardando análise.",
      };
    default:
      return null;
  }
}

function TransitionModal({ processo, config, onClose, onSave }) {
  const [valores, setValores] = useState(() => {
    const iniciais = {};
    config.campos.forEach((c) => { iniciais[c.key] = processo[c.key] || c.default || ""; });
    return iniciais;
  });
  return (
    <ModalShell title={config.titulo} onClose={onClose} maxWidth={440}>
      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
        {config.campos.map((c) => (
          <ModalField key={c.key} label={c.label}>
            <input type={c.type} value={valores[c.key] || ""} onChange={(e) => setValores((v) => ({ ...v, [c.key]: e.target.value }))} style={modalInputStyle} />
          </ModalField>
        ))}
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <button onClick={onClose} style={{ background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "9px 16px", fontSize: 13, cursor: "pointer" }}>Cancelar</button>
        <button onClick={() => { onSave(config.aplicar(valores), config.resumo(valores)); onClose(); }}
          style={{ background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
          {config.textoBotao}
        </button>
      </div>
    </ModalShell>
  );
}

function ConcluirModal({ onClose, onSave }) {
  const [data, setData] = useState(hojeISOStr());
  return (
    <ModalShell title="Marcar Concluído / Deferido" onClose={onClose} maxWidth={400}>
      <ModalField label="Data de conclusão">
        <input type="date" value={data} onChange={(e) => setData(e.target.value)} style={modalInputStyle} />
      </ModalField>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <button onClick={onClose} style={{ background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "9px 16px", fontSize: 13, cursor: "pointer" }}>Cancelar</button>
        <button onClick={() => { onSave({ statusAtual: "concluido", dataConclusao: data }); onClose(); }}
          style={{ background: COLORS.green, border: "none", color: "#0a1420", borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
          Concluído / Deferido
        </button>
      </div>
    </ModalShell>
  );
}

/* ============================================================
   LINHA DO TEMPO — eventos do serviço, do início à conclusão
   ============================================================ */
function eventosLinhaDoTempo(processo) {
  const eventos = [];
  const push = (data, label, cor) => { if (data) eventos.push({ data, label, cor }); };
  const tecnico = processo.tipo === "Serviço Técnico";
  push(processo.dataInicio, "Serviço iniciado", COLORS.blue);
  push(processo.dataPrevisaoAnaliseChecklist, "Previsão de conclusão da análise / checklist", COLORS.steel);
  push(processo.dataPrevistaProtocolo, tecnico ? "Data prevista de vistoria" : "Data prevista de protocolo", COLORS.steel);
  push(processo.dataProtocolo, "Protocolado junto ao órgão", COLORS.blue);
  push(processo.dataPrevisaoOrgao, tecnico ? "Previsão de conclusão / entrega" : "Previsão de análise do órgão", COLORS.steel);
  push(processo.dataExigenciaRecebida, tecnico ? "Pendência registrada" : "Comunique-se / Exigência recebida", COLORS.orange);
  push(processo.dataExigenciaPrazoLimite, "Prazo limite para atendimento", COLORS.red);
  push(processo.dataAtendimentoTecnico, "Atendimento técnico realizado", COLORS.blue);
  push(processo.dataAtendimentoExigencia, "Exigência atendida", COLORS.green);
  (processo.cobrancas || []).forEach((c) => push(c.data, "Cobrança de celeridade ao órgão", COLORS.orange));
  (processo.atualizacoes || []).forEach((a) => push(a.data, `${a.tipo}: ${a.descricao}`, COLORS.steelLight));
  push(processo.dataConclusao, "Concluído / Deferido", COLORS.green);
  return eventos.filter((e) => e.data).sort((a, b) => a.data.localeCompare(b.data));
}

function LinhaDoTempoTab({ processo }) {
  const eventos = eventosLinhaDoTempo(processo);
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
        <div style={{ fontSize: 11.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 600 }}>Linha do tempo do serviço</div>
        <button onClick={() => imprimirLinhaDoTempo(processo)} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "7px 12px", fontSize: 12, cursor: "pointer" }}>
          <Download size={13} /> Exportar linha do tempo
        </button>
      </div>
      {eventos.length === 0 ? (
        <div style={{ fontSize: 12.5, color: COLORS.steel, textAlign: "center", padding: 30 }}>Nenhum evento registrado ainda — inicie o serviço para começar a linha do tempo.</div>
      ) : (
        <div style={{ position: "relative", paddingLeft: 22 }}>
          <div style={{ position: "absolute", left: 5, top: 4, bottom: 4, width: 2, background: COLORS.border }} />
          {eventos.map((e, i) => (
            <div key={i} style={{ position: "relative", paddingBottom: 20 }}>
              <div style={{ position: "absolute", left: -22, top: 2, width: 11, height: 11, borderRadius: "50%", background: e.cor, border: `2px solid ${COLORS.panel}` }} />
              <div style={{ fontSize: 11, color: COLORS.steel, fontFamily: "monospace" }}>{fmtDate(e.data)}</div>
              <div style={{ fontSize: 12.5, color: COLORS.ice, marginTop: 2, lineHeight: 1.4 }}>{e.label}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ============================================================
   EXPORTAÇÕES ELEGANTES — Status de Serviço e Linha do Tempo
   ============================================================ */
const PRINT_BRAND_CSS = `
  * { box-sizing: border-box; }
  body{font-family:'Segoe UI', Arial, Helvetica, sans-serif;color:#16283d;margin:0;background:#fff;}
  .brand{background:#0f1e30;padding:30px 40px;border-bottom:5px solid #e1483d;}
  .brand-row{display:flex;align-items:center;gap:10px;}
  .brand-bars{display:flex;gap:3px;}
  .brand-bars div{width:5px;height:22px;}
  .brand-name{font-family:Georgia,'Times New Roman',serif;font-weight:700;font-size:24px;color:#e1483d;letter-spacing:0.02em;}
  .brand-sub{color:#8493a6;font-size:10.5px;letter-spacing:0.1em;text-transform:uppercase;margin-top:3px;margin-left:1px;}
  .brand-title{color:#eef2f6;font-size:16px;margin-top:18px;font-weight:600;}
  .brand-meta{color:#b7c2cf;font-size:12px;margin-top:4px;}
  .content{padding:30px 40px;}
  h2{font-size:13px;text-transform:uppercase;letter-spacing:.05em;color:#0f1e30;margin:26px 0 10px;border-left:4px solid #e1483d;padding-left:9px;}
  table{width:100%;border-collapse:collapse;font-size:12px;}
  th{background:#0f1e30;color:#eef2f6;padding:9px 10px;text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.03em;}
  td{padding:9px 10px;border-bottom:1px solid #e6e9ed;}
  tr:nth-child(even) td{background:#f7f9fb;}
  .badge{display:inline-block;padding:4px 12px;border-radius:999px;font-size:11px;font-weight:700;}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px 24px;margin-top:10px;}
  .kv{font-size:12px;color:#5b6675;} .kv b{display:block;color:#16283d;font-size:13.5px;font-weight:600;margin-top:2px;}
  .footer{padding:18px 40px;color:#8493a6;font-size:10.5px;border-top:1px solid #e6e9ed;margin-top:20px;}
`;
function brandHeader(title, subtitle) {
  return `<div class="brand">
    <div class="brand-row">
      <div class="brand-bars"><div style="background:#3a4a63"></div><div style="background:#93a2b5"></div><div style="background:#d7dee6"></div><div style="background:#e1483d"></div></div>
      <span class="brand-name">Primers</span>
    </div>
    <div class="brand-sub">Consultoria e Legalização Imobiliária</div>
    <div class="brand-title">${title}</div>
    ${subtitle ? `<div class="brand-meta">${subtitle}</div>` : ""}
  </div>`;
}
function abrirEImprimir(html) {
  const win = window.open("", "_blank");
  if (!win) return;
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 300);
}

function gerarStatusServicoHTML(processo) {
  const status = STATUS_CONFIG[processo.statusAtual];
  const atualizacoes = [...processo.atualizacoes].sort((a, b) => b.data.localeCompare(a.data));
  const linhas = atualizacoes.map((a) => `<tr><td>${fmtDate(a.data)}</td><td>${a.tipo}</td><td>${a.responsavel}</td><td>${a.descricao}</td></tr>`).join("");
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Status de Serviço — ${processo.assunto}</title><style>${PRINT_BRAND_CSS}</style></head><body>
  ${brandHeader("Status de Serviço", `${processo.cliente} — ${processo.unidade} · Gerado em ${fmtDate(hojeISOStr())}`)}
  <div class="content">
    <div class="grid">
      <div class="kv">Serviço<b>${processo.assunto}</b></div>
      <div class="kv">Tipo<b>${processo.tipo}</b></div>
      <div class="kv">Status atual<b><span class="badge" style="background:${status.bg};color:${status.fg}">${status.label}</span></b></div>
      <div class="kv">Nº do processo<b>${processo.numero}</b></div>
      <div class="kv">Data de protocolo<b>${fmtDate(processo.dataProtocolo)}</b></div>
      <div class="kv">Previsão de análise do órgão<b>${fmtDate(processo.dataPrevisaoOrgao)}</b></div>
    </div>
    <h2>Histórico de atualizações</h2>
    <table><tr><th>Data</th><th>Tipo</th><th>Responsável</th><th>Descrição</th></tr>${linhas || `<tr><td colspan="4">Nenhuma atualização registrada.</td></tr>`}</table>
  </div>
  <div class="footer">Primers Consultoria e Legalização Imobiliária — Controle Operacional e Financeiro</div>
  </body></html>`;
}
function imprimirStatusServico(processo) { abrirEImprimir(gerarStatusServicoHTML(processo)); }

function gerarLinhaDoTempoHTML(processo) {
  const eventos = eventosLinhaDoTempo(processo);
  const itens = eventos.map((e) => `
    <div style="display:flex;gap:16px;margin-bottom:18px;">
      <div style="width:90px;flex-shrink:0;font-size:11.5px;color:#5b6675;font-weight:600;padding-top:2px;">${fmtDate(e.data)}</div>
      <div style="flex-shrink:0;padding-top:3px;"><div style="width:11px;height:11px;border-radius:50%;background:${e.cor};"></div></div>
      <div style="font-size:13px;color:#16283d;line-height:1.5;">${e.label}</div>
    </div>`).join("");
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Linha do Tempo — ${processo.assunto}</title><style>${PRINT_BRAND_CSS}</style></head><body>
  ${brandHeader("Linha do Tempo do Serviço", `${processo.cliente} — ${processo.unidade} · ${processo.assunto}`)}
  <div class="content">${itens || `<p>Nenhum evento registrado ainda.</p>`}</div>
  <div class="footer">Primers Consultoria e Legalização Imobiliária — Controle Operacional e Financeiro</div>
  </body></html>`;
}
function imprimirLinhaDoTempo(processo) { abrirEImprimir(gerarLinhaDoTempoHTML(processo)); }

function gerarStatusServicoGeralHTML(processos, tituloCliente) {
  const linhas = processos.map((p) => {
    const st = STATUS_CONFIG[p.statusAtual];
    const ultima = [...p.atualizacoes].sort((a, b) => b.data.localeCompare(a.data))[0];
    return `<tr><td>${p.cliente}</td><td>${p.unidade}</td><td>${p.assunto}</td><td>${p.tipo}</td>
      <td><span class="badge" style="background:${st.bg};color:${st.fg}">${st.label}</span></td>
      <td>${fmtDate(p.ultimaAtualizacao)}</td><td>${ultima ? ultima.descricao : "—"}</td></tr>`;
  }).join("");
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Status de Serviço${tituloCliente ? " — " + tituloCliente : ""}</title><style>${PRINT_BRAND_CSS}</style></head><body>
  ${brandHeader("Status de Serviço", `${tituloCliente || "Todos os clientes"} · Gerado em ${fmtDate(hojeISOStr())} · ${processos.length} serviço(s)`)}
  <div class="content">
    <table><tr><th>Cliente</th><th>Unidade</th><th>Serviço</th><th>Tipo</th><th>Status</th><th>Última atualização</th><th>Última mensagem</th></tr>
    ${linhas || `<tr><td colspan="7">Nenhum serviço encontrado.</td></tr>`}</table>
  </div>
  <div class="footer">Primers Consultoria e Legalização Imobiliária — Controle Operacional e Financeiro</div>
  </body></html>`;
}
function imprimirStatusServicoGeral(processos, tituloCliente) { abrirEImprimir(gerarStatusServicoGeralHTML(processos, tituloCliente)); }

function DetailModal({ processo, processos, onClose, onUpdate, onOpenProcesso, onConcluir }) {
  const [tab, setTab] = useState("geral");
  const [showSenha, setShowSenha] = useState(false);
  const [showAction, setShowAction] = useState(false);
  const [showConcluir, setShowConcluir] = useState(false);
  const status = STATUS_CONFIG[processo.statusAtual];
  const prazo = prazoInfo(diasRestantes(processo));
  const parado = diasSemAtualizacao(processo);
  const bloqueadoPor = processoBloqueado(processo, processos);
  const acao = nextActionConfig(processo);
  const podeConcluir = !status.final;
  const iniciado = processo.statusAtual !== "aguardando";

  const patch = (fields) => onUpdate({ ...processo, ...fields });
  const aplicarTransicao = (fields, resumo) => {
    const hojeISO = hojeISOStr();
    const nova = { id: `transicao-${Date.now()}`, data: hojeISO, tipo: "Tramitação / Movimentação processual", descricao: resumo, responsavel: "Primers" };
    onUpdate({ ...processo, ...fields, ultimaAtualizacao: hojeISO, atualizacoes: [nova, ...processo.atualizacoes] });
  };
  const aplicarConclusao = (fields) => {
    const hojeISO = hojeISOStr();
    const nova = { id: `conclusao-${Date.now()}`, data: fields.dataConclusao || hojeISO, tipo: "Tramitação / Movimentação processual", descricao: "Serviço concluído / deferido.", responsavel: "Primers" };
    const novo = { ...processo, ...fields, ultimaAtualizacao: hojeISO, atualizacoes: [nova, ...processo.atualizacoes] };
    if (onConcluir) onConcluir(novo); else onUpdate(novo);
  };

  const registrarCobranca = () => {
    const nova = { data: hojeISOStr(), nota: "Cobrança de celeridade registrada." };
    patch({ cobrancas: [nova, ...processo.cobrancas] });
  };

  const Row = ({ label, value }) => (
    <div style={{ display: "flex", justifyContent: "space-between", padding: "9px 0", borderBottom: `1px solid ${COLORS.border}` }}>
      <span style={{ fontSize: 12, color: COLORS.steel }}>{label}</span>
      <span style={{ fontSize: 13, color: COLORS.ice, fontWeight: 500, textAlign: "right", maxWidth: "60%" }}>{value}</span>
    </div>
  );

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 40, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ position: "absolute", inset: 0, background: "rgba(5,10,16,0.7)" }} onClick={onClose} />
      <div style={{ position: "relative", width: "100%", maxWidth: 860, maxHeight: "92vh", background: COLORS.panel, border: `1px solid ${COLORS.borderStrong}`, borderRadius: 12, display: "flex", flexDirection: "column", overflow: "hidden" }}>
        {/* HEADER — resumo visível sempre, sem precisar trocar de aba */}
        <div style={{ padding: "18px 22px 14px", borderBottom: `1px solid ${COLORS.border}` }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
            <div>
              <div style={{ fontSize: 11, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em" }}>{processo.cliente} · {processo.unidade}</div>
              <h2 style={{ fontFamily: "'Oswald', sans-serif", fontSize: 18, color: COLORS.ice, fontWeight: 600, marginTop: 3 }}>{processo.assunto}</h2>
            </div>
            <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.steel }}><X size={20} /></button>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 12, flexWrap: "wrap", alignItems: "center" }}>
            <select value={processo.statusAtual} onChange={(e) => patch({ statusAtual: e.target.value })}
              style={{ background: status.bg, color: status.fg, border: `1px solid ${status.fg}55`, borderRadius: 999, padding: "4px 10px", fontSize: 11, fontWeight: 700, fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em" }}>
              {STATUS_KEYS.map((k) => <option key={k} value={k} style={{ background: COLORS.panel, color: COLORS.ice }}>{STATUS_CONFIG[k].label}</option>)}
            </select>
            <Pill fg={COLORS.steelLight} bg="rgba(255,255,255,0.06)">{processo.tipo}</Pill>
            {iniciado && <Pill fg={prazo.fg} bg={prazo.bg}>{prazo.label}</Pill>}
            {iniciado && parado !== null && <Pill fg={parado > 15 ? COLORS.red : COLORS.steel} bg={parado > 15 ? COLORS.redDim : "rgba(255,255,255,0.05)"}>{parado > 15 ? `Parado há ${parado}d` : `Atualizado há ${parado}d`}</Pill>}
            {iniciado && <Pill fg={checklistProgress(processo.checklist) === 100 ? COLORS.green : COLORS.steelLight} bg="rgba(255,255,255,0.06)">Checklist {checklistProgress(processo.checklist)}%</Pill>}
            {iniciado && <Pill fg={documentosProgress(processo.documentos) === 100 ? COLORS.green : COLORS.steelLight} bg="rgba(255,255,255,0.06)">Documentos {documentosProgress(processo.documentos)}%</Pill>}
            {processo.pendenciaCliente.ativa && <Pill fg={COLORS.yellow} bg={COLORS.yellowDim}>Pendência do cliente</Pill>}
            {bloqueadoPor && <Pill fg={COLORS.red} bg={COLORS.redDim}>Bloqueado por outro processo</Pill>}
          </div>

          {iniciado && (
            <div style={{ display: "flex", gap: 4, marginTop: 14, flexWrap: "wrap" }}>
              {[
                ["geral", "Visão geral", Building2],
                ["documentos", "Documentos", FileStack],
                ["urbanistico", "Parâmetros urbanísticos", Building2],
                ["checklist", "Checklist", ClipboardCheck],
                ["linhadotempo", "Linha do tempo", Timer],
                ["atualizacoes", `Status de Serviço (${processo.atualizacoes.length})`, History],
                ["relatorio", "Relatório", Download],
              ].map(([id, label, Icon]) => (
                <button key={id} onClick={() => setTab(id)} style={{
                  display: "flex", alignItems: "center", gap: 6, background: tab === id ? COLORS.redDim : "transparent",
                  color: tab === id ? COLORS.red : COLORS.steelLight, border: "none", borderRadius: 7, padding: "7px 12px", fontSize: 12.5, fontWeight: tab === id ? 700 : 500, cursor: "pointer",
                }}><Icon size={13} />{label}</button>
              ))}
            </div>
          )}

          {(acao || podeConcluir) && (
            <div style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap", alignItems: "center", background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 12px" }}>
              <span style={{ fontSize: 11, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.03em", fontWeight: 700 }}>Próximo passo:</span>
              {acao && (
                <button onClick={() => setShowAction(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "7px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.02em", textTransform: "uppercase" }}>
                  {acao.textoBotao}
                </button>
              )}
              {podeConcluir && processo.statusAtual !== "aguardando" && (
                <button onClick={() => setShowConcluir(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLORS.green}55`, color: COLORS.green, borderRadius: 7, padding: "7px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>
                  <CheckCircle2 size={13} /> Concluído / Deferido
                </button>
              )}
            </div>
          )}
        </div>

        {/* BODY */}
        <div style={{ padding: "18px 22px", overflowY: "auto", flex: 1 }}>
          {!iniciado && (
            <div style={{ textAlign: "center", padding: "50px 20px" }}>
              <Clock size={34} color={COLORS.steel} style={{ marginBottom: 14 }} />
              <div style={{ fontSize: 15, color: COLORS.ice, fontFamily: "'Oswald', sans-serif", fontWeight: 600, marginBottom: 6 }}>Este serviço ainda não foi iniciado</div>
              <div style={{ fontSize: 12.5, color: COLORS.steel, marginBottom: 20, maxWidth: 380, marginLeft: "auto", marginRight: "auto" }}>
                Clique em "Iniciar Serviço" para registrar a data de início e a previsão de conclusão da análise documental / checklist. Documentos, Checklist, Linha do tempo e Status de Serviço ficam disponíveis depois disso.
              </div>
              <button onClick={() => setShowAction(true)} style={{ background: COLORS.red, border: "none", color: "#fff", borderRadius: 8, padding: "11px 24px", fontSize: 13.5, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
                Iniciar Serviço
              </button>
              <div style={{ marginTop: 24, textAlign: "left", maxWidth: 380, marginLeft: "auto", marginRight: "auto" }}>
                <Row label="Cliente / Unidade" value={`${processo.cliente} — ${processo.unidade}`} />
                <Row label="Tipo de serviço" value={processo.tipo} />
                <Row label="Técnico" value={processo.tecnico} />
                <Row label="Contrato vinculado" value={processo.numeroContrato} />
              </div>
            </div>
          )}

          {iniciado && tab === "geral" && (
            <div>
              <Row label="Cidade / UF" value={`${processo.cidade} - ${processo.uf}`} />
              <Row label="Tipo de serviço" value={processo.tipo} />
              <Row label="Técnico" value={processo.tecnico} />
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: `1px solid ${COLORS.border}` }}>
                <span style={{ fontSize: 12, color: COLORS.steel }}>Nº do processo</span>
                <input value={processo.numero === "-" ? "" : processo.numero} onChange={(e) => patch({ numero: e.target.value || "-" })} placeholder="Preencher quando obtido"
                  style={{ background: "transparent", border: "none", borderBottom: `1px dashed ${COLORS.border}`, color: COLORS.ice, fontSize: 13, textAlign: "right", padding: "2px 0" }} />
              </div>
              <Row label="Data de início" value={fmtDate(processo.dataInicio)} />
              <Row label="Data de protocolo" value={fmtDate(processo.dataProtocolo)} />
              <Row label="Previsão de análise do órgão" value={fmtDate(processo.dataPrevisaoOrgao)} />
              <Row label="Exigência recebida" value={fmtDate(processo.dataExigenciaRecebida)} />
              <Row label="Prazo limite para atendimento" value={fmtDate(processo.dataExigenciaPrazoLimite)} />
              <Row label="Exigência atendida" value={fmtDate(processo.dataAtendimentoExigencia)} />
              <Row label="Data de conclusão" value={fmtDate(processo.dataConclusao)} />
              <Row label="Prestador" value={processo.prestador || "—"} />
              <Row label="Site do órgão" value={processo.site || "—"} />
              <Row label="Login" value={processo.login || "—"} />
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: `1px solid ${COLORS.border}` }}>
                <span style={{ fontSize: 12, color: COLORS.steel }}>Senha</span>
                <button onClick={() => setShowSenha((s) => !s)} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", cursor: processo.senha && processo.senha !== "-" ? "pointer" : "default", color: COLORS.ice, fontSize: 13, fontFamily: showSenha ? "'Inter', sans-serif" : "monospace" }}>
                  {processo.senha && processo.senha !== "-" ? (showSenha ? processo.senha : "••••••••••") : "—"}
                  {processo.senha && processo.senha !== "-" && (showSenha ? <EyeOff size={14} color={COLORS.steel} /> : <Eye size={14} color={COLORS.steel} />)}
                </button>
              </div>

              <div style={{ marginTop: 16, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ fontSize: 11.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 600 }}>Cobranças de celeridade</div>
                <button onClick={registrarCobranca} style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 6, padding: "5px 10px", fontSize: 11.5, cursor: "pointer" }}>
                  <Plus size={12} /> Registrar cobrança
                </button>
              </div>
              {processo.cobrancas.length === 0 ? (
                <div style={{ fontSize: 12, color: COLORS.steel, padding: "10px 0" }}>Nenhuma cobrança registrada.</div>
              ) : processo.cobrancas.map((c, i) => (
                <div key={i} style={{ fontSize: 12, color: COLORS.steelLight, padding: "6px 0", borderBottom: `1px solid ${COLORS.border}` }}>{fmtDate(c.data)} — {c.nota}</div>
              ))}

              <div style={{ marginTop: 16 }}>
                <div style={{ fontSize: 11.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 600, marginBottom: 8 }}>Pendência do cliente</div>
                <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, cursor: "pointer" }}>
                  <input type="checkbox" checked={processo.pendenciaCliente.ativa} onChange={(e) => patch({ pendenciaCliente: { ...processo.pendenciaCliente, ativa: e.target.checked } })} />
                  <span style={{ fontSize: 12.5, color: COLORS.ice }}>Há pendência aberta do cliente</span>
                </label>
                {processo.pendenciaCliente.ativa && (
                  <input value={processo.pendenciaCliente.descricao} onChange={(e) => patch({ pendenciaCliente: { ...processo.pendenciaCliente, descricao: e.target.value } })} placeholder="Descreva a pendência..."
                    style={{ width: "100%", background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "8px 10px", color: COLORS.ice, fontSize: 12.5 }} />
                )}
              </div>

              <div style={{ marginTop: 20, fontSize: 11.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 600, marginBottom: 8 }}>Referência do contrato</div>
              <div style={{ fontSize: 12.5, color: COLORS.steelLight, background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 12px", lineHeight: 1.6 }}>
                Contrato <b style={{ color: COLORS.ice }}>{processo.numeroContrato}</b> · Valor do serviço <b style={{ color: COLORS.ice }}>{fmtBRL(processo.valorContrato)}</b>.
                O status de faturamento (Faturado / Concluído-Não faturado / Em andamento / Suspenso) é controlado na tela de Contratos e no Planejamento Financeiro.
              </div>

              <div style={{ marginTop: 20, fontSize: 11.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 600, marginBottom: 8 }}>Dependência de outro processo</div>
              <select value={processo.dependeDeId || ""} onChange={(e) => patch({ dependeDeId: e.target.value ? Number(e.target.value) : null })}
                style={{ width: "100%", background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "7px 9px", color: COLORS.ice, fontSize: 12.5 }}>
                <option value="">Nenhuma dependência</option>
                {processos.filter((p) => p.id !== processo.id).map((p) => <option key={p.id} value={p.id}>{p.cliente} — {p.assunto}</option>)}
              </select>
              {bloqueadoPor && (
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 8, background: COLORS.redDim, border: `1px solid ${COLORS.red}55`, borderRadius: 8, padding: "9px 12px" }}>
                  <div style={{ fontSize: 12, color: COLORS.red }}>Bloqueado até a conclusão de: <b>{bloqueadoPor.assunto}</b> ({bloqueadoPor.cliente})</div>
                  {onOpenProcesso && <button onClick={() => onOpenProcesso(bloqueadoPor)} style={{ background: "transparent", border: "none", color: COLORS.red, fontSize: 12, cursor: "pointer", textDecoration: "underline" }}>Abrir</button>}
                </div>
              )}
            </div>
          )}
          {iniciado && tab === "documentos" && <DocumentosTab processo={processo} onUpdate={onUpdate} />}
          {iniciado && tab === "urbanistico" && <ParametrosUrbanisticosTab processo={processo} onUpdate={onUpdate} />}
          {iniciado && tab === "checklist" && <ChecklistTab processo={processo} onUpdate={onUpdate} />}
          {iniciado && tab === "linhadotempo" && <LinhaDoTempoTab processo={processo} />}
          {iniciado && tab === "atualizacoes" && <AtualizacoesTab processo={processo} onUpdate={onUpdate} />}
          {iniciado && tab === "relatorio" && <RelatorioTab processo={processo} />}
        </div>
      </div>
      {showAction && acao && <TransitionModal processo={processo} config={acao} onClose={() => setShowAction(false)} onSave={aplicarTransicao} />}
      {showConcluir && <ConcluirModal onClose={() => setShowConcluir(false)} onSave={aplicarConclusao} />}
    </div>
  );
}

/* ============================================================
   IMPORT CSV MODAL
   ============================================================ */
function ImportModal({ onClose, onImport }) {
  const [status, setStatus] = useState(null);
  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      Papa.parse(ev.target.result, {
        header: true, skipEmptyLines: true,
        complete: (res) => {
          const mapped = res.data.map((row) => baseProcesso({
            id: Date.now() + Math.random(),
            cliente: row["Cliente"] || "", unidade: row["Unidade"] || "", cidade: row["Cidade"] || "", uf: row["UF"] || "",
            assunto: row["Assunto / Serviço"] || row["Assunto"] || "", tipo: row["Tipo / Serviço"] || "Processo",
            numero: row["Nº Processo / Documento"] || "-", statusAtual: "protocolado",
            dataProtocolo: row["Data prot. / emissão"] || null, dataPrevisaoOrgao: row["Data conclusão prevista"] || null,
            site: row["Site"] || "-", login: row["Login"] || "-", senha: row["Senha"] || "-", prestador: row["Prestador"] || "-",
          })).filter((r) => r.cliente);
          setStatus({ ok: true, count: mapped.length, data: mapped });
        },
        error: () => setStatus({ ok: false }),
      });
    };
    reader.readAsText(file, "UTF-8");
  };
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(5,10,16,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }} onClick={onClose}>
      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.borderStrong}`, borderRadius: 12, width: "100%", maxWidth: 480, padding: 24 }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <h2 style={{ fontFamily: "'Oswald', sans-serif", fontSize: 18, color: COLORS.ice, fontWeight: 600, textTransform: "uppercase" }}>Importar planilha (CSV)</h2>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.steel }}><X size={20} /></button>
        </div>
        <p style={{ fontSize: 12.5, color: COLORS.steel, lineHeight: 1.6, marginBottom: 16 }}>
          Exporte sua planilha como <b style={{ color: COLORS.steelLight }}>CSV</b> mantendo os cabeçalhos originais e envie o arquivo abaixo. Os processos importados entram como "Protocolado — aguardando análise"; ajuste o status depois, se necessário.
        </p>
        <input type="file" accept=".csv" onChange={handleFile} style={{ color: COLORS.steelLight, fontSize: 12.5 }} />
        {status && status.ok && <div style={{ marginTop: 14, fontSize: 13, color: COLORS.green }}>{status.count} processo(s) reconhecido(s).</div>}
        {status && !status.ok && <div style={{ marginTop: 14, fontSize: 13, color: COLORS.red }}>Não foi possível ler esse arquivo.</div>}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
          <button onClick={onClose} style={{ background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "9px 16px", fontSize: 13, cursor: "pointer" }}>Cancelar</button>
          <button disabled={!status || !status.ok} onClick={() => { onImport(status.data); onClose(); }}
            style={{ background: status && status.ok ? COLORS.red : COLORS.grayDim, border: "none", color: status && status.ok ? "#fff" : COLORS.steel, borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 700, cursor: status && status.ok ? "pointer" : "default", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
            Confirmar importação
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   PÁGINA: ATUALIZAÇÕES (feed global + exportação)
   ============================================================ */
/* ============================================================
   AGENDA SEMANAL — prazos do Controle de Processos + itens
   avulsos cadastrados manualmente (reuniões, tarefas, lembretes)
   ============================================================ */
const AGENDA_TIPO_OPTIONS = ["Reunião", "Tarefa", "Lembrete", "Outro"];
const AGENDA_TIPO_COLOR = { "Reunião": COLORS.blue, "Tarefa": COLORS.orange, "Lembrete": COLORS.yellow, "Outro": COLORS.steelLight };

function AgendaItemModal({ dataInicial, onClose, onSave }) {
  const [f, setF] = useState({ data: dataInicial || new Date().toISOString().slice(0, 10), titulo: "", tipo: "Reunião", tecnico: "", descricao: "" });
  const set = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));
  return (
    <ModalShell title="Novo item da agenda" onClose={onClose} maxWidth={440}>
      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
        <ModalField label="Data"><input type="date" value={f.data} onChange={set("data")} style={modalInputStyle} /></ModalField>
        <ModalField label="Título"><input value={f.titulo} onChange={set("titulo")} placeholder="Ex: Reunião com cliente X" style={modalInputStyle} /></ModalField>
        <ModalField label="Tipo">
          <select value={f.tipo} onChange={set("tipo")} style={modalInputStyle}>
            {AGENDA_TIPO_OPTIONS.map((o) => <option key={o} value={o}>{o}</option>)}
          </select>
        </ModalField>
        <ModalField label="Técnico (opcional)">
          <select value={f.tecnico} onChange={set("tecnico")} style={modalInputStyle}>
            <option value="">Não atribuído</option>
            {TECNICOS_OPTIONS.map((o) => <option key={o} value={o}>{o}</option>)}
          </select>
        </ModalField>
        <ModalField label="Descrição (opcional)"><input value={f.descricao} onChange={set("descricao")} placeholder="Detalhes, participantes, local..." style={modalInputStyle} /></ModalField>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <button onClick={onClose} style={{ background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "9px 16px", fontSize: 13, cursor: "pointer" }}>Cancelar</button>
        <button onClick={() => {
          if (!f.titulo.trim()) return;
          onSave({ id: `agenda-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`, data: f.data, titulo: f.titulo.trim(), tipo: f.tipo, tecnico: f.tecnico, descricao: f.descricao.trim() });
          onClose();
        }} style={{ background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
          Salvar item
        </button>
      </div>
    </ModalShell>
  );
}

function compromissosDoProcesso(p) {
  const tecnico = p.tipo === "Serviço Técnico";
  const lista = [];
  if (p.dataPrevisaoAnaliseChecklist) lista.push({ data: p.dataPrevisaoAnaliseChecklist, label: "Previsão de análise/checklist" });
  if (p.dataPrevistaProtocolo) lista.push({ data: p.dataPrevistaProtocolo, label: tecnico ? "Previsão de vistoria" : "Previsão de protocolo" });
  if (p.dataPrevisaoOrgao) lista.push({ data: p.dataPrevisaoOrgao, label: tecnico ? "Previsão de entrega" : "Previsão de análise do órgão" });
  if (p.dataExigenciaPrazoLimite) lista.push({ data: p.dataExigenciaPrazoLimite, label: "Prazo limite da exigência" });
  return lista;
}

function AgendaSemanal({ processos, agendaItens, onOpenProcesso, onAddItem, onRemoveItem }) {
  const [showModal, setShowModal] = useState(null); // data (iso) do dia clicado, ou null
  const [tecnicoAtivo, setTecnicoAtivo] = useState("Todos");
  const hojeRef = new Date();
  const diaSemana = hojeRef.getDay();
  const diffSegunda = diaSemana === 0 ? -6 : 1 - diaSemana;
  const segunda = new Date(hojeRef);
  segunda.setDate(hojeRef.getDate() + diffSegunda);
  segunda.setHours(0, 0, 0, 0);
  const nomesDia = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"];
  const isoDia = (d) => d.toISOString().slice(0, 10);
  const hojeIso = isoDia(hojeRef);

  const processosFiltrados = tecnicoAtivo === "Todos" ? processos : processos.filter((p) => p.tecnico === tecnicoAtivo);
  const itensFiltrados = tecnicoAtivo === "Todos" ? agendaItens : agendaItens.filter((a) => !a.tecnico || a.tecnico === tecnicoAtivo);

  const compromissosPorDia = {};
  processosFiltrados.forEach((p) => {
    compromissosDoProcesso(p).forEach((c) => {
      if (!compromissosPorDia[c.data]) compromissosPorDia[c.data] = [];
      compromissosPorDia[c.data].push({ processo: p, label: c.label });
    });
  });

  const dias = Array.from({ length: 7 }, (_, i) => {
    const d = new Date(segunda);
    d.setDate(segunda.getDate() + i);
    const iso = isoDia(d);
    return {
      label: nomesDia[i], data: d, iso,
      processosDia: compromissosPorDia[iso] || [],
      itensDia: itensFiltrados.filter((a) => a.data === iso),
    };
  });

  return (
    <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 18, marginBottom: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
        <div style={{ fontSize: 12, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600, display: "flex", alignItems: "center", gap: 7 }}>
          <Clock size={13} /> Agenda da semana — prazos e itens cadastrados
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <div style={{ display: "flex", gap: 4 }}>
            {["Todos", ...TECNICOS_OPTIONS].map((t) => (
              <button key={t} onClick={() => setTecnicoAtivo(t)} style={{
                background: tecnicoAtivo === t ? COLORS.redDim : "transparent", color: tecnicoAtivo === t ? COLORS.red : COLORS.steelLight,
                border: `1px solid ${tecnicoAtivo === t ? COLORS.red + "55" : COLORS.border}`, borderRadius: 6, padding: "5px 10px", fontSize: 11.5, fontWeight: tecnicoAtivo === t ? 700 : 500, cursor: "pointer",
              }}>{t}</button>
            ))}
          </div>
          <button onClick={() => setShowModal(hojeIso)} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 6, padding: "5px 10px", fontSize: 11.5, cursor: "pointer" }}>
            <Plus size={12} /> Novo item
          </button>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, minmax(120px, 1fr))", gap: 10, overflowX: "auto" }}>
        {dias.map(({ label, data, iso, processosDia, itensDia }) => (
          <div key={iso} style={{
            background: iso === hojeIso ? COLORS.redDim : COLORS.panelAlt,
            border: `1px solid ${iso === hojeIso ? COLORS.red + "55" : COLORS.border}`, borderRadius: 8, padding: 10, minHeight: 140, display: "flex", flexDirection: "column",
          }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ fontSize: 10.5, color: iso === hojeIso ? COLORS.red : COLORS.steel, textTransform: "uppercase", fontWeight: 700, marginBottom: 2 }}>{label}</div>
                <div style={{ fontSize: 15, color: COLORS.ice, fontFamily: "'Oswald', sans-serif", fontWeight: 600, marginBottom: 6 }}>
                  {String(data.getDate()).padStart(2, "0")}/{String(data.getMonth() + 1).padStart(2, "0")}
                </div>
              </div>
              <button onClick={() => setShowModal(iso)} title="Adicionar item neste dia"
                style={{ background: "transparent", border: `1px solid ${COLORS.border}`, borderRadius: 5, width: 20, height: 20, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 }}>
                <Plus size={11} color={COLORS.steel} />
              </button>
            </div>

            {processosDia.slice(0, 2).map((c, ci) => (
              <div key={ci} onClick={() => onOpenProcesso(c.processo)} style={{ fontSize: 10.5, color: COLORS.steelLight, marginBottom: 5, cursor: "pointer", lineHeight: 1.4 }}>
                <div style={{ fontWeight: 600, color: COLORS.ice }}>{c.processo.cliente}</div>
                <div>{c.processo.assunto}</div>
                <div style={{ fontSize: 9.5, color: COLORS.steel }}>{c.label}</div>
              </div>
            ))}
            {processosDia.length > 2 && <div style={{ fontSize: 10, color: COLORS.steel, marginBottom: 4 }}>+{processosDia.length - 2} processo(s)</div>}

            {itensDia.map((a) => (
              <div key={a.id} style={{ display: "flex", alignItems: "flex-start", gap: 5, marginBottom: 5 }}>
                <div style={{ width: 6, height: 6, borderRadius: "50%", background: AGENDA_TIPO_COLOR[a.tipo] || COLORS.steelLight, marginTop: 4, flexShrink: 0 }} />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 10.5, color: COLORS.ice, fontWeight: 600, lineHeight: 1.3 }}>{a.titulo}</div>
                  <div style={{ fontSize: 9.5, color: COLORS.steel }}>{a.tipo}{a.tecnico ? ` · ${a.tecnico}` : ""}{a.descricao ? ` · ${a.descricao}` : ""}</div>
                </div>
                <button onClick={() => onRemoveItem(a.id)} title="Remover" style={{ background: "none", border: "none", cursor: "pointer", padding: 0, flexShrink: 0 }}>
                  <X size={10} color={COLORS.steel} />
                </button>
              </div>
            ))}

            {processosDia.length === 0 && itensDia.length === 0 && <div style={{ fontSize: 10.5, color: COLORS.steel }}>—</div>}
          </div>
        ))}
      </div>
      {showModal && <AgendaItemModal dataInicial={showModal} onClose={() => setShowModal(null)} onSave={onAddItem} />}
    </div>
  );
}

function AtualizacoesPage({ processos, onOpenProcesso }) {
  const [filtroCliente, setFiltroCliente] = useState("Todos");
  const [filtroTipo, setFiltroTipo] = useState("Todos");
  const [filtroTecnico, setFiltroTecnico] = useState("Todos");
  const [exportCliente, setExportCliente] = useState("Todos");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);

  const clientes = useMemo(() => ["Todos", ...Array.from(new Set(processos.map((p) => p.cliente))).sort()], [processos]);

  const feed = useMemo(() => {
    const items = [];
    processos.forEach((p) => {
      if (filtroCliente !== "Todos" && p.cliente !== filtroCliente) return;
      if (filtroTecnico !== "Todos" && p.tecnico !== filtroTecnico) return;
      p.atualizacoes.forEach((a) => {
        if (filtroTipo !== "Todos" && a.tipo !== filtroTipo) return;
        items.push({ ...a, processo: p });
      });
    });
    return items.sort((a, b) => new Date(b.data) - new Date(a.data));
  }, [processos, filtroCliente, filtroTipo, filtroTecnico]);
  const paginado = useMemo(() => paginate(feed, page, pageSize), [feed, page, pageSize]);

  const exportar = () => {
    const alvo = exportCliente === "Todos" ? processos : processos.filter((p) => p.cliente === exportCliente);
    const rows = [["Cliente", "Unidade", "Assunto", "Status atual", "Responsável", "Data protocolo", "Previsão órgão", "Última atualização", "Última mensagem"]];
    alvo.forEach((p) => {
      const ultima = p.atualizacoes[0];
      rows.push([
        p.cliente, p.unidade, p.assunto, STATUS_CONFIG[p.statusAtual].label, STATUS_CONFIG[p.statusAtual].responsavel,
        fmtDate(p.dataProtocolo), fmtDate(p.dataPrevisaoOrgao), fmtDate(p.ultimaAtualizacao), ultima ? ultima.descricao : "",
      ]);
    });
    downloadCSV(`situacao_processos_${exportCliente === "Todos" ? "geral" : exportCliente}.csv`, rows);
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 18, justifyContent: "space-between" }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <Select value={filtroCliente} onChange={(v) => { setFiltroCliente(v); setPage(1); }} options={clientes.filter((c) => c !== "Todos")} placeholder="Todos os clientes" />
          <Select value={filtroTecnico} onChange={(v) => { setFiltroTecnico(v); setPage(1); }} options={TECNICOS_OPTIONS} placeholder="Todos os técnicos" />
          <Select value={filtroTipo} onChange={(v) => { setFiltroTipo(v); setPage(1); }} options={ATUALIZACAO_TIPOS} placeholder="Todos os tipos" />
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <Select value={exportCliente} onChange={setExportCliente} options={clientes.filter((c) => c !== "Todos")} placeholder="Exportar: todos os clientes" />
          <button onClick={() => imprimirStatusServicoGeral(exportCliente === "Todos" ? processos : processos.filter((p) => p.cliente === exportCliente), exportCliente === "Todos" ? null : exportCliente)}
            style={{ display: "flex", alignItems: "center", gap: 6, background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "8px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.02em", textTransform: "uppercase" }}>
            <Download size={13} /> Exportar para o cliente
          </button>
          <button onClick={exportar} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
            <Download size={13} /> CSV
          </button>
        </div>
      </div>

      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: "4px 0" }}>
        {feed.length === 0 && <div style={{ padding: 30, textAlign: "center", color: COLORS.steel, fontSize: 13 }}>Nenhuma atualização encontrada com esses filtros.</div>}
        {paginado.map((a) => (
          <div key={a.id} className="row-hover" onClick={() => onOpenProcesso(a.processo)} style={{ display: "flex", gap: 14, padding: "12px 18px", borderBottom: `1px solid ${COLORS.border}`, cursor: "pointer" }}>
            <div style={{ width: 78, flexShrink: 0, fontSize: 11.5, color: COLORS.steel, fontFamily: "monospace" }}>{fmtDate(a.data)}</div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 3 }}>
                <span style={{ fontSize: 13, fontWeight: 600, color: COLORS.ice }}>{a.processo.cliente}</span>
                <span style={{ fontSize: 11.5, color: COLORS.steel }}>{a.processo.unidade}</span>
                <Pill fg={COLORS.steelLight} bg="rgba(255,255,255,0.06)">{a.tipo}</Pill>
                <span style={{ fontSize: 10.5, color: COLORS.steel }}>resp.: {a.responsavel}</span>
              </div>
              <div style={{ fontSize: 11.5, color: COLORS.steel, marginBottom: 3 }}>{a.processo.assunto} · técnico: {a.processo.tecnico}</div>
              <div style={{ fontSize: 12.5, color: COLORS.steelLight, lineHeight: 1.5 }}>{a.descricao}</div>
            </div>
            <ChevronRight size={15} color={COLORS.steel} style={{ flexShrink: 0, marginTop: 2 }} />
          </div>
        ))}
        <Pagination page={page} setPage={setPage} pageSize={pageSize} setPageSize={setPageSize} totalItems={feed.length} />
      </div>
    </div>
  );
}

/* ============================================================
   CONTRATOS — página financeira/comercial alimentada por upload
   de planilha (ex: planejamento-financeiro), independente dos
   processos de produção
   ============================================================ */
const CONTRATO_HEADER_MAP = {
  proposta: ["Proposta"],
  cliente: ["Cliente"],
  unidade: ["Unidade"],
  codigoLoja: ["Código Loja", "Codigo Loja", "Código  Loja"],
  servico: ["Serviço", "Servico"],
  tarefa: ["Tarefa"],
  tecnico: ["Técnico", "Tecnico"],
  coordenador: ["Coordenador"],
  honorarios: ["Honorários", "Honorarios"],
  valorFaturamento: ["Valor Faturamento"],
  porcentagem: ["Porcentagem"],
  dataSLA: ["Data SLA"],
  dataFaturamento: ["Data Faturamento"],
  dataSLAServico: ["Data SLA Serviço", "Data SLA Servico"],
  statusContrato: ["Status do Contrato"],
  statusServico: ["Status do Serviço", "Status do Servico"],
  statusParcela: ["Status da Parcela"],
  observacao: ["Observação", "Observacao"],
};

function findHeaderIndex(headerRow, aliases) {
  const norm = (s) => String(s || "").trim().toLowerCase();
  for (const alias of aliases) {
    const idx = headerRow.findIndex((h) => norm(h) === norm(alias));
    if (idx !== -1) return idx;
  }
  return -1;
}

function excelDateToISO(v) {
  if (v === null || v === undefined || v === "" || v === "-") return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "number") {
    const d = XLSX.SSF ? XLSX.SSF.parse_date_code(v) : null;
    if (d) return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  }
  const str = String(v).trim();
  const m1 = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m1) return `${m1[1]}-${m1[2]}-${m1[3]}`;
  const m2 = str.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m2) return `${m2[3]}-${m2[2]}-${m2[1]}`;
  return null;
}

function parseContratosSheet(rows) {
  if (!rows || rows.length < 2) return [];
  const header = rows[0];
  const cols = {};
  Object.entries(CONTRATO_HEADER_MAP).forEach(([key, aliases]) => { cols[key] = findHeaderIndex(header, aliases); });
  const get = (row, key) => (cols[key] !== -1 ? row[cols[key]] : undefined);
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;
    const cliente = get(row, "cliente");
    if (!cliente || String(cliente).trim() === "") continue;
    out.push({
      id: `${Date.now()}-${i}-${Math.random().toString(36).slice(2, 7)}`,
      proposta: String(get(row, "proposta") ?? "-").trim() || "-",
      cliente: String(cliente).trim(),
      unidade: String(get(row, "unidade") ?? "-").trim() || "-",
      codigoLoja: String(get(row, "codigoLoja") ?? "-").trim() || "-",
      servico: String(get(row, "servico") ?? "-").trim() || "-",
      tarefa: String(get(row, "tarefa") ?? "-").trim() || "-",
      tecnico: String(get(row, "tecnico") ?? "-").trim() || "-",
      coordenador: String(get(row, "coordenador") ?? "-").trim() || "-",
      tipo: classifyTipoServico(get(row, "servico")),
      honorarios: parseFloat(get(row, "honorarios")) || 0,
      valorFaturamento: parseFloat(get(row, "valorFaturamento")) || 0,
      porcentagem: parseFloat(get(row, "porcentagem")) || 0,
      dataSLA: excelDateToISO(get(row, "dataSLA")),
      dataFaturamento: excelDateToISO(get(row, "dataFaturamento")),
      dataSLAServico: excelDateToISO(get(row, "dataSLAServico")),
      statusContrato: String(get(row, "statusContrato") ?? "-").trim() || "-",
      statusServico: String(get(row, "statusServico") ?? "-").trim() || "-",
      statusParcela: normalizeStatusParcela(get(row, "statusParcela")),
      observacao: String(get(row, "observacao") ?? "").trim(),
    });
  }
  return out;
}

const STATUS_PARCELA_OPTIONS = ["Concluído / Não faturado", "Em andamento", "Faturado", "Suspenso"];
function normalizeStatusParcela(raw) {
  const s = String(raw || "").trim().toLowerCase();
  if (s === "invoice" || s === "faturado" || s === "pago") return "Faturado";
  if (s === "em progresso" || s === "em andamento" || s === "andamento") return "Em andamento";
  if (s === "suspenso") return "Suspenso";
  return "Concluído / Não faturado";
}
const CONTRATO_TIPO_OPTIONS = ["Processo", "Serviço Técnico"];
const TECNICOS_OPTIONS = ["Felipe Moura", "Janayna"];

/* ============================================================
   VÍNCULO CONTRATOS → PROCESSOS
   Cada linha da planilha é uma parcela/tarefa de um serviço.
   Agrupamos por proposta+cliente+unidade+serviço para formar (ou
   atualizar) um processo em Produção, com data e atualizações
   derivadas — assim Processos, Atualizações e os Dashboards
   passam a refletir automaticamente o que foi importado.
   ============================================================ */
function classifyTipoServico(servico) {
  const s = (servico || "").toLowerCase();
  const tecnico = ["art/rrt", "art ", " art", "rrt", "atestado", "manifestação técnica", "manifestacao tecnica",
    "as built", "as-built", "vistoria", "memorial", "relatório de análises", "relatorio de analises",
    "transferência de responsabilidade", "transferencia de responsabilidade"];
  return tecnico.some((k) => s.includes(k)) ? "Serviço Técnico" : "Processo";
}

function mapContratoStatus(statusContrato, statusServico) {
  const c = (statusContrato || "").trim();
  const s = (statusServico || "").trim();
  if (/conclu[ií]do/i.test(c)) return "concluido";
  if (/cancelado/i.test(c)) return "cancelado";
  if (/suspenso/i.test(c) || /suspenso/i.test(s)) return "suspenso";
  if (s === "Distribuido") return "em_montagem";
  if (s === "Faturar") return "protocolado";
  if (s === "Pendente") return "em_montagem";
  if (s === "Em Andamento") return "protocolado";
  if (/pendente/i.test(c)) return "em_montagem";
  return "protocolado";
}

function deriveStatusFaturamento(parcelas) {
  const algumFaturado = parcelas.some((p) => p.statusParcela === "Faturado" || p.dataFaturamento);
  return algumFaturado ? "faturado" : "nao_faturado";
}

function minDate(dates) { const v = dates.filter(Boolean).sort(); return v[0] || null; }
function maxDate(dates) { const v = dates.filter(Boolean).sort(); return v[v.length - 1] || null; }

function tarefaParaTipoAtualizacao(tarefa) {
  const t = (tarefa || "").toLowerCase();
  if (t.includes("protocolo")) return "Tramitação / Movimentação processual";
  if (t.includes("deferimento") || t.includes("obtenção") || t.includes("entrega")) return "Tramitação / Movimentação processual";
  if (t.includes("exigência") || t.includes("exigencia") || t.includes("nec")) return "Comunique-se / Exigência recebida";
  if (t.includes("sinal") || t.includes("contratação")) return "Reunião / Alinhamento com cliente";
  return "Outro";
}

/* Ao importar uma nova planilha, concilia com os contratos já
   existentes: atualiza a linha correspondente (mesma proposta +
   cliente + unidade + serviço + tarefa) em vez de duplicar, e
   mantém tudo que já estava cadastrado e não veio na nova planilha. */
function mergeImportedContratos(existentes, novasLinhas) {
  const chave = (c) => `${c.proposta}|${c.cliente}|${c.unidade}|${c.servico}|${c.tarefa}`;
  const out = [...existentes];
  novasLinhas.forEach((nova) => {
    const idx = out.findIndex((c) => chave(c) === chave(nova));
    if (idx !== -1) out[idx] = { ...out[idx], ...nova, id: out[idx].id };
    else out.push(nova);
  });
  return out;
}

function grupoContratos(rows) {
  const map = {};
  rows.forEach((r) => {
    const key = `${r.proposta}|${r.cliente}|${r.unidade}|${r.servico}`;
    if (!map[key]) map[key] = [];
    map[key].push(r);
  });
  return map;
}

function mergeContratosIntoProcessos(processosAtuais, novasLinhas) {
  const grupos = grupoContratos(novasLinhas);
  let processos = [...processosAtuais];

  Object.values(grupos).forEach((parcelas) => {
    const ref = parcelas[0];
    const idxExistente = processos.findIndex((p) => p.numeroContrato === ref.proposta && p.cliente === ref.cliente && p.unidade === ref.unidade && p.assunto === ref.servico);
    const statusFaturamento = deriveStatusFaturamento(parcelas);
    const honorarios = Math.max(0, ...parcelas.map((p) => p.honorarios || 0));
    const dataFaturamento = maxDate(parcelas.map((p) => p.dataFaturamento));
    const hojeISO = new Date().toISOString().slice(0, 10);

    if (idxExistente !== -1) {
      // Processo já existe e pode estar em andamento manualmente — a reimportação só
      // atualiza o vínculo financeiro (valor, faturamento, tipo), nunca mexe no status
      // operacional, nas datas ou nas atualizações já registradas manualmente.
      const atual = processos[idxExistente];
      processos[idxExistente] = {
        ...atual,
        statusFaturamento, valorContrato: honorarios, tipo: ref.tipo || atual.tipo,
        tecnico: TECNICOS_OPTIONS.includes(ref.tecnico) ? ref.tecnico : atual.tecnico,
        dataFaturamento: dataFaturamento || atual.dataFaturamento,
        numeroContrato: ref.proposta,
        parcelasContrato: parcelas,
      };
    } else {
      processos = [baseProcesso({
        id: `contrato-${ref.proposta}-${ref.servico}-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        cliente: ref.cliente, unidade: ref.unidade, cidade: "", uf: "",
        assunto: ref.servico, tipo: ref.tipo || classifyTipoServico(ref.servico), numero: "-",
        statusAtual: "aguardando", prestador: "-", tecnico: TECNICOS_OPTIONS.includes(ref.tecnico) ? ref.tecnico : "-",
        numeroContrato: ref.proposta, valorContrato: honorarios, statusFaturamento, dataFaturamento,
        parcelasContrato: parcelas, ultimaAtualizacao: hojeISO,
      }), ...processos];
    }
  });

  return processos;
}

/* Quando um processo é concluído/deferido, a(s) parcela(s) do contrato
   vinculado passam para "Concluído / Não faturado" (nunca rebaixa uma
   parcela que já esteja Faturada). */
function sincronizarParcelasNaConclusao(processo, contratosAtuais) {
  return contratosAtuais.map((c) => {
    if (c.proposta === processo.numeroContrato && c.cliente === processo.cliente && c.unidade === processo.unidade && c.servico === processo.assunto && c.statusParcela !== "Faturado") {
      return { ...c, statusParcela: "Concluído / Não faturado" };
    }
    return c;
  });
}

/* O técnico (e o tipo) só podem ser alterados na tela de Contratos —
   qualquer edição ali reflete automaticamente no processo vinculado. */
function sincronizarTecnicoTipoProcesso(contrato, processosAtuais) {
  return processosAtuais.map((p) => {
    if (p.numeroContrato === contrato.proposta && p.cliente === contrato.cliente && p.unidade === contrato.unidade && p.assunto === contrato.servico) {
      return { ...p, tecnico: TECNICOS_OPTIONS.includes(contrato.tecnico) ? contrato.tecnico : p.tecnico, tipo: contrato.tipo || p.tipo };
    }
    return p;
  });
}

const STATUS_PARCELA_COLOR = {
  "Concluído / Não faturado": { fg: COLORS.orange, bg: COLORS.orangeDim },
  "Em andamento": { fg: COLORS.blue, bg: COLORS.blueDim },
  "Faturado": { fg: COLORS.green, bg: COLORS.greenDim },
  "Suspenso": { fg: COLORS.steel, bg: COLORS.grayDim },
};
function statusParcelaStyle(s) { return STATUS_PARCELA_COLOR[s] || { fg: COLORS.steelLight, bg: "rgba(255,255,255,0.06)" }; }
const STATUS_CONTRATO_COLOR = {
  "Em Andamento": { fg: COLORS.blue, bg: COLORS.blueDim },
  "Pendente": { fg: COLORS.steel, bg: "rgba(255,255,255,0.06)" },
  "Suspenso": { fg: COLORS.orange, bg: COLORS.orangeDim },
  "Concluído": { fg: COLORS.green, bg: COLORS.greenDim },
  "Cancelado": { fg: COLORS.overdue, bg: COLORS.overdueDim },
};
function statusContratoStyle(s) { return STATUS_CONTRATO_COLOR[s] || { fg: COLORS.steelLight, bg: "rgba(255,255,255,0.06)" }; }

/* ============================================================
   PAGINAÇÃO — reutilizada em todas as listagens (clientes,
   unidades, contratos, processos, atualizações)
   ============================================================ */
const PAGE_SIZE_OPTIONS = [10, 20, 50, 100];
function paginate(items, page, pageSize) {
  if (pageSize === "todas") return items;
  const start = (page - 1) * pageSize;
  return items.slice(start, start + pageSize);
}
function Pagination({ page, setPage, pageSize, setPageSize, totalItems }) {
  const totalPages = pageSize === "todas" ? 1 : Math.max(1, Math.ceil(totalItems / pageSize));
  const from = totalItems === 0 ? 0 : (pageSize === "todas" ? 1 : (page - 1) * pageSize + 1);
  const to = pageSize === "todas" ? totalItems : Math.min(page * pageSize, totalItems);
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 16px", flexWrap: "wrap", gap: 10, borderTop: `1px solid ${COLORS.border}` }}>
      <div style={{ fontSize: 11.5, color: COLORS.steel }}>{totalItems === 0 ? "Nenhum registro" : `${from}–${to} de ${totalItems}`}</div>
      <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ fontSize: 11.5, color: COLORS.steel }}>Linhas por página</span>
          <select value={pageSize} onChange={(e) => { setPageSize(e.target.value === "todas" ? "todas" : Number(e.target.value)); setPage(1); }}
            style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "5px 8px", color: COLORS.steelLight, fontSize: 12 }}>
            {PAGE_SIZE_OPTIONS.map((n) => <option key={n} value={n}>{n}</option>)}
            <option value="todas">Todas</option>
          </select>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <button disabled={page <= 1} onClick={() => setPage((p) => Math.max(1, p - 1))}
            style={{ background: "transparent", border: `1px solid ${COLORS.border}`, borderRadius: 6, width: 26, height: 26, display: "flex", alignItems: "center", justifyContent: "center", cursor: page <= 1 ? "default" : "pointer", opacity: page <= 1 ? 0.4 : 1 }}>
            <ChevronLeft size={14} color={COLORS.steelLight} />
          </button>
          <span style={{ fontSize: 11.5, color: COLORS.steelLight, minWidth: 90, textAlign: "center" }}>Página {page} de {totalPages}</span>
          <button disabled={page >= totalPages} onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
            style={{ background: "transparent", border: `1px solid ${COLORS.border}`, borderRadius: 6, width: 26, height: 26, display: "flex", alignItems: "center", justifyContent: "center", cursor: page >= totalPages ? "default" : "pointer", opacity: page >= totalPages ? 0.4 : 1 }}>
            <ChevronRight size={14} color={COLORS.steelLight} />
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   IMPORTAR NOVOS CLIENTES / CONTRATOS — página dedicada
   (único lugar onde a planilha de contratos é enviada)
   ============================================================ */
function ImportarClientesContratosPage({ onImport }) {
  const [status, setStatus] = useState(null);
  const [confirmado, setConfirmado] = useState(false);

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setConfirmado(false);
    const isCSV = /\.csv$/i.test(file.name);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        let rows;
        if (isCSV) {
          const parsed = Papa.parse(ev.target.result, { skipEmptyLines: true });
          rows = parsed.data;
        } else {
          const wb = XLSX.read(ev.target.result, { type: "array", cellDates: true });
          const sheetName = wb.SheetNames.find((n) => /planejamento|contrato|financeiro/i.test(n)) || wb.SheetNames[0];
          const ws = wb.Sheets[sheetName];
          rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
        }
        const contratos = parseContratosSheet(rows);
        const clientesNovos = new Set(contratos.map((c) => c.cliente)).size;
        const unidadesNovas = new Set(contratos.map((c) => `${c.cliente}|${c.unidade}`)).size;
        setStatus({ ok: true, count: contratos.length, clientesNovos, unidadesNovas, data: contratos, fileName: file.name });
      } catch (err) {
        setStatus({ ok: false });
      }
    };
    if (isCSV) reader.readAsText(file, "UTF-8");
    else reader.readAsArrayBuffer(file);
  };

  return (
    <div style={{ maxWidth: 620 }}>
      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 24 }}>
        <div style={{ fontSize: 11.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700, marginBottom: 6 }}>Importar planilha</div>
        <p style={{ fontSize: 12.5, color: COLORS.steel, lineHeight: 1.6, marginBottom: 18 }}>
          Envie a planilha de contratos (.xlsx) no mesmo formato do planejamento financeiro — com colunas como Proposta, Cliente, Unidade, Serviço, Tarefa, Honorários, Valor Faturamento, datas e status.
          Os clientes, unidades e contratos são adicionados aos já existentes, e os processos e atualizações correspondentes são criados ou atualizados automaticamente.
        </p>
        <label style={{
          display: "flex", flexDirection: "column", alignItems: "center", gap: 8, border: `1.5px dashed ${COLORS.border}`,
          borderRadius: 10, padding: "28px 16px", cursor: "pointer", background: COLORS.panelAlt,
        }}>
          <Upload size={22} color={COLORS.steel} />
          <span style={{ fontSize: 13, color: COLORS.steelLight, fontWeight: 600 }}>Clique para selecionar o arquivo</span>
          <span style={{ fontSize: 11, color: COLORS.steel }}>.xlsx, .xls ou .csv</span>
          <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
        </label>

        {status && status.ok && (
          <div style={{ marginTop: 18, background: COLORS.greenDim, border: `1px solid ${COLORS.green}55`, borderRadius: 8, padding: 14 }}>
            <div style={{ fontSize: 13, color: COLORS.green, fontWeight: 700, marginBottom: 4 }}>{status.fileName} lido com sucesso</div>
            <div style={{ fontSize: 12.5, color: COLORS.steelLight, lineHeight: 1.6 }}>
              {status.count} linha(s) de contrato · {status.clientesNovos} cliente(s) · {status.unidadesNovas} unidade(s) reconhecida(s).
            </div>
          </div>
        )}
        {status && !status.ok && (
          <div style={{ marginTop: 18, fontSize: 13, color: COLORS.red }}>Não foi possível ler esse arquivo. Verifique se é um .xlsx ou .csv válido com os cabeçalhos esperados.</div>
        )}
        {confirmado && <div style={{ marginTop: 14, fontSize: 13, color: COLORS.green }}>Importação concluída — dados disponíveis em Clientes, Unidades, Contratos, Processos e Atualizações.</div>}

        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 20 }}>
          <button disabled={!status || !status.ok} onClick={() => { onImport(status.data); setConfirmado(true); setStatus(null); }}
            style={{ background: status && status.ok ? COLORS.red : COLORS.grayDim, border: "none", color: status && status.ok ? "#fff" : COLORS.steel, borderRadius: 7, padding: "10px 20px", fontSize: 13, fontWeight: 700, cursor: status && status.ok ? "pointer" : "default", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
            Confirmar importação
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   CLIENTES — listagem agregada a partir dos contratos importados
   ============================================================ */
/* ============================================================
   CADASTRO MANUAL — cliente / unidade / contrato
   (alternativa à importação de planilha)
   ============================================================ */
function novaLinhaContrato(f) {
  return {
    id: `manual-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    proposta: f.proposta || "-", cliente: f.cliente, unidade: f.unidade || "-", codigoLoja: f.codigoLoja || "-",
    servico: f.servico || "-", tarefa: f.tarefa || "-", tecnico: f.tecnico || "-", coordenador: f.coordenador || "-",
    tipo: f.tipo || classifyTipoServico(f.servico),
    honorarios: f.honorarios || 0, valorFaturamento: f.valorFaturamento || 0, porcentagem: f.porcentagem || 0,
    dataSLA: f.dataSLA || null, dataFaturamento: f.dataFaturamento || null, dataSLAServico: null,
    statusContrato: f.statusContrato || "Em Andamento", statusServico: f.statusServico || "Em Andamento",
    statusParcela: f.statusParcela || "Concluído / Não faturado", observacao: f.observacao || "",
  };
}

/* ============================================================
   LOGIN — autenticação real via Supabase (sem senha no código)
   ============================================================ */
/* ============================================================
   GERENCIAR ACESSOS — só o administrador vê esta tela. Cria e
   remove logins chamando o backend seguro (Edge Function), que
   guarda a chave secreta do lado do servidor.
   ============================================================ */
function GerenciarAcessosPage({ usuarioLogado }) {
  const [usuarios, setUsuarios] = useState([]);
  const [carregando, setCarregando] = useState(true);
  const [erro, setErro] = useState("");
  const [form, setForm] = useState({ usuario: "", senha: "", nome: "", role: "operacional" });
  const [salvando, setSalvando] = useState(false);

  const carregar = async () => {
    setCarregando(true);
    const resp = await callAdminUsers("list");
    if (resp.error) setErro(resp.error); else setUsuarios(resp.usuarios || []);
    setCarregando(false);
  };
  useEffect(() => { carregar(); }, []);

  const criar = async () => {
    if (!form.usuario.trim() || !form.senha || !form.nome.trim()) { setErro("Preencha usuário, senha e nome."); return; }
    setSalvando(true); setErro("");
    const resp = await callAdminUsers("create", { usuario: form.usuario.trim().toLowerCase(), senha: form.senha, nome: form.nome.trim(), role: form.role });
    setSalvando(false);
    if (resp.error) { setErro(resp.error); return; }
    setForm({ usuario: "", senha: "", nome: "", role: "operacional" });
    carregar();
  };

  const remover = async (u) => {
    if (!window.confirm(`Remover o acesso de "${u.nome}" (${u.usuario})? Essa ação não pode ser desfeita.`)) return;
    const resp = await callAdminUsers("delete", { id: u.id });
    if (resp.error) setErro(resp.error); else carregar();
  };

  return (
    <div>
      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 20, marginBottom: 20, maxWidth: 560 }}>
        <div style={{ fontSize: 11.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700, marginBottom: 14 }}>Novo acesso</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
          <ModalField label="Usuário (login)">
            <input value={form.usuario} onChange={(e) => setForm((f) => ({ ...f, usuario: e.target.value }))} placeholder="ex: janayna" style={modalInputStyle} />
          </ModalField>
          <ModalField label="Senha">
            <input type="password" value={form.senha} onChange={(e) => setForm((f) => ({ ...f, senha: e.target.value }))} placeholder="mínimo 6 caracteres" style={modalInputStyle} />
          </ModalField>
          <ModalField label="Nome">
            <input value={form.nome} onChange={(e) => setForm((f) => ({ ...f, nome: e.target.value }))} placeholder="ex: Janayna" style={modalInputStyle} />
          </ModalField>
          <ModalField label="Papel">
            <select value={form.role} onChange={(e) => setForm((f) => ({ ...f, role: e.target.value }))} style={modalInputStyle}>
              <option value="operacional">Operacional</option>
              <option value="admin">Administrador</option>
            </select>
          </ModalField>
        </div>
        {erro && <div style={{ fontSize: 12, color: COLORS.red, marginBottom: 10 }}>{erro}</div>}
        <button onClick={criar} disabled={salvando} style={{ background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 700, cursor: salvando ? "default" : "pointer", opacity: salvando ? 0.7 : 1, fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
          {salvando ? "Criando..." : "Criar acesso"}
        </button>
      </div>

      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, overflow: "hidden" }}>
        <div style={{ fontSize: 11.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700, padding: "14px 18px 10px" }}>Acessos cadastrados</div>
        <table>
          <thead><tr>{["Usuário", "Nome", "Papel", ""].map((h) => (
            <th key={h} style={{ textAlign: "left", padding: "8px 18px", fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${COLORS.border}` }}>{h}</th>
          ))}</tr></thead>
          <tbody>
            {usuarios.map((u) => (
              <tr key={u.id} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                <td style={{ padding: "10px 18px", fontSize: 13, fontFamily: "monospace", color: COLORS.steelLight }}>{u.usuario}</td>
                <td style={{ padding: "10px 18px", fontSize: 13, color: COLORS.ice, fontWeight: 600 }}>{u.nome}</td>
                <td style={{ padding: "10px 18px" }}><Pill fg={u.role === "admin" ? COLORS.red : COLORS.steelLight} bg={u.role === "admin" ? COLORS.redDim : "rgba(255,255,255,0.06)"}>{u.role === "admin" ? "Administrador" : "Operacional"}</Pill></td>
                <td style={{ padding: "10px 18px", textAlign: "right" }}>
                  {u.id !== usuarioLogado.id && (
                    <button onClick={() => remover(u)} style={{ background: "transparent", border: `1px solid ${COLORS.red}55`, borderRadius: 6, width: 26, height: 26, display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer" }}>
                      <Trash2 size={12} color={COLORS.red} />
                    </button>
                  )}
                </td>
              </tr>
            ))}
            {!carregando && usuarios.length === 0 && <tr><td colSpan={4} style={{ padding: 24, textAlign: "center", color: COLORS.steel, fontSize: 13 }}>Nenhum acesso encontrado.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function LoginScreen({ onLogin }) {
  const [usuario, setUsuario] = useState("");
  const [senha, setSenha] = useState("");
  const [erro, setErro] = useState("");
  const [carregando, setCarregando] = useState(false);

  const entrar = async () => {
    if (!SUPABASE_CONFIGURADO) { setErro("O sistema ainda não está conectado ao banco de dados (Partes 5 e 6 do guia). Complete essas partes e tente de novo."); return; }
    if (!usuario.trim() || !senha) { setErro("Preencha usuário e senha."); return; }
    setCarregando(true);
    setErro("");
    const emailInterno = `${usuario.trim().toLowerCase()}@primers.local`;
    const { data, error } = await supabase.auth.signInWithPassword({ email: emailInterno, password: senha });
    if (error || !data.user) {
      setCarregando(false);
      setErro("Login ou senha inválidos.");
      return;
    }
    const { data: perfil, error: perfilErro } = await supabase.from("profiles").select("id, usuario, nome, role").eq("id", data.user.id).single();
    setCarregando(false);
    if (perfilErro || !perfil) { setErro("Login sem perfil cadastrado. Fale com o administrador."); return; }
    onLogin(perfil);
  };

  return (
    <div style={{ minHeight: "100vh", background: COLORS.bg, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif", padding: 16 }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap'); * { box-sizing: border-box; } ::placeholder { color: ${COLORS.steel}; opacity: 0.7; }`}</style>
      <div style={{ width: "100%", maxWidth: 360, background: COLORS.panel, border: `1px solid ${COLORS.borderStrong}`, borderRadius: 12, padding: 32 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 3, marginBottom: 4 }}>
          <div style={{ width: 5, height: 22, background: "#3a4a63" }} />
          <div style={{ width: 5, height: 22, background: "#93a2b5" }} />
          <div style={{ width: 5, height: 22, background: "#d7dee6" }} />
          <div style={{ width: 5, height: 22, background: COLORS.red }} />
          <span style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 700, fontSize: 19, color: COLORS.red, marginLeft: 8, letterSpacing: "0.02em" }}>PRIMERS</span>
        </div>
        <div style={{ fontSize: 10.5, color: COLORS.steel, letterSpacing: "0.06em", marginBottom: 28 }}>CONTROL — acesso restrito</div>

        {!SUPABASE_CONFIGURADO && (
          <div style={{ background: COLORS.yellowDim, border: `1px solid ${COLORS.yellow}55`, borderRadius: 8, padding: "10px 12px", marginBottom: 20, fontSize: 11.5, color: COLORS.yellow, lineHeight: 1.5 }}>
            Backend ainda não conectado. Complete as Partes 5 e 6 do guia (criar o Supabase e colar a URL/chave no início deste arquivo) para poder entrar.
          </div>
        )}

        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <ModalField label="Usuário">
            <input value={usuario} onChange={(e) => setUsuario(e.target.value)} onKeyDown={(e) => e.key === "Enter" && entrar()} placeholder="felipe ou janayna" style={modalInputStyle} autoFocus />
          </ModalField>
          <ModalField label="Senha">
            <input type="password" value={senha} onChange={(e) => setSenha(e.target.value)} onKeyDown={(e) => e.key === "Enter" && entrar()} placeholder="••••••" style={modalInputStyle} />

          </ModalField>
          {erro && <div style={{ fontSize: 12, color: COLORS.red }}>{erro}</div>}
          <button onClick={entrar} disabled={carregando} style={{ marginTop: 6, background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "11px 18px", fontSize: 13, fontWeight: 700, cursor: carregando ? "default" : "pointer", opacity: carregando ? 0.7 : 1, fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
            {carregando ? "Entrando..." : "Entrar"}
          </button>
        </div>
        <div style={{ fontSize: 11, color: COLORS.steel, marginTop: 20, lineHeight: 1.5 }}>
          Protótipo de testes internos — Primers Consultoria e Legalização Imobiliária.
        </div>
      </div>
    </div>
  );
}

function ModalShell({ title, onClose, maxWidth = 460, children }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(5,10,16,0.7)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }} onClick={onClose}>
      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.borderStrong}`, borderRadius: 12, width: "100%", maxWidth, maxHeight: "88vh", overflowY: "auto", padding: 24 }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
          <h2 style={{ fontFamily: "'Oswald', sans-serif", fontSize: 18, fontWeight: 600, color: COLORS.ice, textTransform: "uppercase", letterSpacing: "0.03em" }}>{title}</h2>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.steel }}><X size={20} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}
function ModalField({ label, children }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      <label style={{ fontSize: 11, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 600 }}>{label}</label>
      {children}
    </div>
  );
}
const modalInputStyle = { background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "8px 10px", color: COLORS.ice, fontSize: 13, fontFamily: "'Inter', sans-serif", outline: "none", width: "100%" };

function NovoClienteModal({ onClose, onSave }) {
  const [cliente, setCliente] = useState("");
  return (
    <ModalShell title="Novo cliente" onClose={onClose}>
      <ModalField label="Nome do cliente">
        <input value={cliente} onChange={(e) => setCliente(e.target.value)} placeholder="Ex: Atacado Boreal" style={modalInputStyle} />
      </ModalField>
      <p style={{ fontSize: 11.5, color: COLORS.steel, marginTop: 10, lineHeight: 1.5 }}>
        O cliente entra na listagem sem contrato vinculado ainda. Cadastre um contrato depois (em Contratos) para registrar serviços e valores.
      </p>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 18 }}>
        <button onClick={onClose} style={{ background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "9px 16px", fontSize: 13, cursor: "pointer" }}>Cancelar</button>
        <button onClick={() => { if (!cliente.trim()) return; onSave(novaLinhaContrato({ cliente: cliente.trim(), servico: "Cadastro manual (sem contrato ainda)", statusContrato: "-", statusServico: "-", statusParcela: "Concluído / Não faturado" })); onClose(); }}
          style={{ background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
          Salvar cliente
        </button>
      </div>
    </ModalShell>
  );
}

function NovaUnidadeModal({ onClose, onSave, clientesExistentes }) {
  const [cliente, setCliente] = useState("");
  const [unidade, setUnidade] = useState("");
  return (
    <ModalShell title="Nova unidade" onClose={onClose}>
      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
        <ModalField label="Cliente">
          <input value={cliente} onChange={(e) => setCliente(e.target.value)} placeholder="Digite ou escolha um cliente existente" list="clientes-existentes" style={modalInputStyle} />
          <datalist id="clientes-existentes">{clientesExistentes.map((c) => <option key={c} value={c} />)}</datalist>
        </ModalField>
        <ModalField label="Unidade">
          <input value={unidade} onChange={(e) => setUnidade(e.target.value)} placeholder="Ex: Loja Centro" style={modalInputStyle} />
        </ModalField>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 18 }}>
        <button onClick={onClose} style={{ background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "9px 16px", fontSize: 13, cursor: "pointer" }}>Cancelar</button>
        <button onClick={() => { if (!cliente.trim() || !unidade.trim()) return; onSave(novaLinhaContrato({ cliente: cliente.trim(), unidade: unidade.trim(), servico: "Cadastro manual (sem contrato ainda)", statusContrato: "-", statusServico: "-", statusParcela: "Concluído / Não faturado" })); onClose(); }}
          style={{ background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
          Salvar unidade
        </button>
      </div>
    </ModalShell>
  );
}

function ContratoFormModal({ title, submitLabel, initial, onClose, onSubmit, clientesExistentes, isAdmin }) {
  const [f, setF] = useState({
    proposta: "", cliente: "", unidade: "", codigoLoja: "", servico: "", tarefa: "", tecnico: "", coordenador: "",
    tipo: "Processo", honorarios: "", valorFaturamento: "", porcentagemPct: "100", dataSLA: "", statusParcela: "Concluído / Não faturado", observacao: "",
    ...initial,
  });
  const set = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));
  const field = (label, key, placeholder, type = "text", extra) => (
    <ModalField label={label}><input type={type} value={f[key]} onChange={set(key)} placeholder={placeholder} style={modalInputStyle} {...extra} /></ModalField>
  );

  return (
    <ModalShell title={title} onClose={onClose} maxWidth={640}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        {field("Proposta", "proposta", "Ex: PC-05200-08-26")}
        <ModalField label="Cliente">
          <input value={f.cliente} onChange={set("cliente")} placeholder="Digite ou escolha um cliente existente" list="clientes-existentes-contrato" style={modalInputStyle} />
          <datalist id="clientes-existentes-contrato">{clientesExistentes.map((c) => <option key={c} value={c} />)}</datalist>
        </ModalField>
        {field("Unidade", "unidade", "Ex: Loja Centro")}
        {field("Código loja", "codigoLoja", "Ex: 12")}
        <div style={{ gridColumn: "1 / -1" }}>{field("Serviço", "servico", "Ex: Aprovação de Publicidade junto a Prefeitura")}</div>
        {field("Tarefa / parcela", "tarefa", "Ex: Sinal, Protocolo, Deferimento, Entrega")}
        <ModalField label="Tipo">
          <select value={f.tipo} onChange={set("tipo")} style={modalInputStyle}>
            {CONTRATO_TIPO_OPTIONS.map((o) => <option key={o} value={o}>{o}</option>)}
          </select>
        </ModalField>
        <ModalField label="Técnico">
          <select value={TECNICOS_OPTIONS.includes(f.tecnico) ? f.tecnico : ""} onChange={set("tecnico")} style={modalInputStyle}>
            <option value="">Selecionar...</option>
            {TECNICOS_OPTIONS.map((o) => <option key={o} value={o}>{o}</option>)}
          </select>
        </ModalField>
        {field("Coordenador", "coordenador", "Nome do coordenador")}
        {field("Data SLA", "dataSLA", "", "date")}
        {isAdmin ? (
          <>
            {field("Honorários (R$)", "honorarios", "Valor total do serviço", "number")}
            {field("Valor da parcela (R$)", "valorFaturamento", "Valor desta parcela", "number")}
          </>
        ) : (
          <div style={{ gridColumn: "1 / -1", fontSize: 11.5, color: COLORS.steel, background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "9px 12px" }}>
            Valores visíveis apenas para administradores.
          </div>
        )}
        {field("Porcentagem da parcela (%)", "porcentagemPct", "Ex: 50", "number")}
        <ModalField label="Status da parcela">
          <select value={f.statusParcela} onChange={set("statusParcela")} style={modalInputStyle}>
            {STATUS_PARCELA_OPTIONS.map((o) => <option key={o} value={o}>{o}</option>)}
          </select>
        </ModalField>
        <div style={{ gridColumn: "1 / -1" }}>{field("Observação", "observacao", "Opcional")}</div>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <button onClick={onClose} style={{ background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "9px 16px", fontSize: 13, cursor: "pointer" }}>Cancelar</button>
        <button onClick={() => {
          if (!f.cliente.trim() || !f.servico.trim()) return;
          onSubmit({
            proposta: f.proposta.trim(), cliente: f.cliente.trim(), unidade: f.unidade.trim(), codigoLoja: f.codigoLoja.trim(),
            servico: f.servico.trim(), tarefa: f.tarefa.trim(), tecnico: f.tecnico.trim(), coordenador: f.coordenador.trim(), tipo: f.tipo,
            honorarios: parseFloat(f.honorarios) || 0, valorFaturamento: parseFloat(f.valorFaturamento) || 0,
            porcentagem: (parseFloat(f.porcentagemPct) || 0) / 100, dataSLA: f.dataSLA || null,
            statusParcela: f.statusParcela, observacao: f.observacao.trim(),
          });
          onClose();
        }} style={{ background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
          {submitLabel}
        </button>
      </div>
    </ModalShell>
  );
}

function ClientesPage({ contratos, onAddContrato, isAdmin }) {
  const [busca, setBusca] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);

  const clientes = useMemo(() => {
    const map = {};
    contratos.forEach((c) => {
      if (!map[c.cliente]) map[c.cliente] = { cliente: c.cliente, unidades: new Set(), propostas: new Set() };
      const m = map[c.cliente];
      m.unidades.add(c.unidade);
      m.propostas.add(c.proposta);
    });
    return Object.values(map).map((m) => ({ cliente: m.cliente, unidades: m.unidades.size, propostas: m.propostas.size })).sort((a, b) => a.cliente.localeCompare(b.cliente));
  }, [contratos]);

  const filtrados = useMemo(() => clientes.filter((c) => !busca || c.cliente.toLowerCase().includes(busca.toLowerCase())), [clientes, busca]);
  const paginados = useMemo(() => paginate(filtrados, page, pageSize), [filtrados, page, pageSize]);
  const [showNovo, setShowNovo] = useState(false);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "8px 12px", maxWidth: 360, flex: 1 }}>
          <Search size={14} color={COLORS.steel} />
          <input value={busca} onChange={(e) => { setBusca(e.target.value); setPage(1); }} placeholder="Buscar cliente..."
            style={{ background: "transparent", border: "none", outline: "none", color: COLORS.ice, fontSize: 13, width: "100%" }} />
        </div>
        {isAdmin && (
          <button onClick={() => setShowNovo(true)} style={{ display: "flex", alignItems: "center", gap: 7, background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.02em", textTransform: "uppercase" }}>
            <Plus size={15} /> Novo cliente
          </button>
        )}
      </div>
      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr>{["Cliente", "Unidades", "Contratos"].map((h) => (
              <th key={h} style={{ textAlign: "left", padding: "10px 16px", fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${COLORS.border}` }}>{h}</th>
            ))}</tr></thead>
            <tbody>
              {paginados.map((c) => (
                <tr key={c.cliente} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                  <td style={{ padding: "11px 16px", fontSize: 13, fontWeight: 600, color: COLORS.ice, display: "flex", alignItems: "center", gap: 6 }}><Building2 size={12} color={COLORS.steel} />{c.cliente}</td>
                  <td style={{ padding: "11px 16px", fontSize: 13, color: COLORS.steelLight }}>{c.unidades}</td>
                  <td style={{ padding: "11px 16px", fontSize: 13, color: COLORS.steelLight }}>{c.propostas}</td>
                </tr>
              ))}
              {filtrados.length === 0 && <tr><td colSpan={3} style={{ padding: 40, textAlign: "center", color: COLORS.steel, fontSize: 13 }}>Nenhum cliente importado ainda.</td></tr>}
            </tbody>
          </table>
        </div>
        <Pagination page={page} setPage={setPage} pageSize={pageSize} setPageSize={setPageSize} totalItems={filtrados.length} />
      </div>
      {showNovo && isAdmin && <NovoClienteModal onClose={() => setShowNovo(false)} onSave={onAddContrato} />}
    </div>
  );
}

/* ============================================================
   UNIDADES DE CLIENTES — listagem agregada por cliente + unidade
   ============================================================ */
function UnidadesPage({ contratos, onAddContrato }) {
  const [busca, setBusca] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);

  const unidades = useMemo(() => {
    const map = {};
    contratos.forEach((c) => {
      const key = `${c.cliente}|${c.unidade}`;
      if (!map[key]) map[key] = { cliente: c.cliente, unidade: c.unidade, servicos: new Set(), statusContrato: c.statusContrato };
      const m = map[key];
      m.servicos.add(c.servico);
      m.statusContrato = c.statusContrato || m.statusContrato;
    });
    return Object.values(map).map((m) => ({ cliente: m.cliente, unidade: m.unidade, servicos: m.servicos.size, statusContrato: m.statusContrato })).sort((a, b) => a.cliente.localeCompare(b.cliente) || a.unidade.localeCompare(b.unidade));
  }, [contratos]);

  const filtrados = useMemo(() => unidades.filter((u) => !busca || `${u.cliente} ${u.unidade}`.toLowerCase().includes(busca.toLowerCase())), [unidades, busca]);
  const paginados = useMemo(() => paginate(filtrados, page, pageSize), [filtrados, page, pageSize]);
  const clientesExistentes = useMemo(() => Array.from(new Set(contratos.map((c) => c.cliente))).sort(), [contratos]);
  const [showNovo, setShowNovo] = useState(false);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "8px 12px", maxWidth: 360, flex: 1 }}>
          <Search size={14} color={COLORS.steel} />
          <input value={busca} onChange={(e) => { setBusca(e.target.value); setPage(1); }} placeholder="Buscar cliente ou unidade..."
            style={{ background: "transparent", border: "none", outline: "none", color: COLORS.ice, fontSize: 13, width: "100%" }} />
        </div>
        <button onClick={() => setShowNovo(true)} style={{ display: "flex", alignItems: "center", gap: 7, background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.02em", textTransform: "uppercase" }}>
          <Plus size={15} /> Nova unidade
        </button>
      </div>
      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr>{["Cliente", "Unidade", "Serviços", "Status do contrato"].map((h) => (
              <th key={h} style={{ textAlign: "left", padding: "10px 16px", fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${COLORS.border}` }}>{h}</th>
            ))}</tr></thead>
            <tbody>
              {paginados.map((u, i) => {
                const sc = statusContratoStyle(u.statusContrato);
                return (
                  <tr key={i} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                    <td style={{ padding: "11px 16px", fontSize: 13, fontWeight: 600, color: COLORS.ice }}>{u.cliente}</td>
                    <td style={{ padding: "11px 16px", fontSize: 13, color: COLORS.steelLight }}>{u.unidade}</td>
                    <td style={{ padding: "11px 16px", fontSize: 13, color: COLORS.steelLight }}>{u.servicos}</td>
                    <td style={{ padding: "11px 16px" }}><Pill fg={sc.fg} bg={sc.bg}>{u.statusContrato}</Pill></td>
                  </tr>
                );
              })}
              {filtrados.length === 0 && <tr><td colSpan={4} style={{ padding: 40, textAlign: "center", color: COLORS.steel, fontSize: 13 }}>Nenhuma unidade importada ainda.</td></tr>}
            </tbody>
          </table>
        </div>
        <Pagination page={page} setPage={setPage} pageSize={pageSize} setPageSize={setPageSize} totalItems={filtrados.length} />
      </div>
      {showNovo && <NovaUnidadeModal onClose={() => setShowNovo(false)} onSave={onAddContrato} clientesExistentes={clientesExistentes} />}
    </div>
  );
}

/* ============================================================
   PLANEJAMENTO MENSAL — visão do financeiro por mês/ano,
   conforme a Data SLA de cada parcela do contrato
   ============================================================ */
function PlanejamentoMensalPage({ contratos }) {
  const anosDisponiveis = useMemo(() => {
    const anos = new Set(contratos.map((c) => (c.dataSLA ? c.dataSLA.slice(0, 4) : null)).filter(Boolean));
    if (anos.size === 0) anos.add(String(new Date().getFullYear()));
    return Array.from(anos).sort();
  }, [contratos]);
  const [ano, setAno] = useState(anosDisponiveis[anosDisponiveis.length - 1]);
  const [mes, setMes] = useState("Todos");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);
  useEffect(() => { if (!anosDisponiveis.includes(ano)) setAno(anosDisponiveis[anosDisponiveis.length - 1]); }, [anosDisponiveis]); // eslint-disable-line


  const doAno = useMemo(() => contratos.filter((c) => c.dataSLA && c.dataSLA.slice(0, 4) === ano), [contratos, ano]);
  const filtrados = useMemo(() => {
    if (mes === "Todos") return doAno;
    return doAno.filter((c) => String(parseInt(c.dataSLA.slice(5, 7), 10)) === mes);
  }, [doAno, mes]);

  const totalPlanejado = filtrados.reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const totalFaturado = filtrados.filter((c) => c.statusParcela === "Faturado").reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const totalAndamento = filtrados.filter((c) => c.statusParcela === "Em andamento").reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const totalConcluidoNaoFaturado = filtrados.filter((c) => c.statusParcela === "Concluído / Não faturado").reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const totalSuspenso = filtrados.filter((c) => c.statusParcela === "Suspenso").reduce((s, c) => s + (c.valorFaturamento || 0), 0);

  const chartData = useMemo(() => {
    if (mes === "Todos") {
      const map = {};
      for (let m = 1; m <= 12; m++) map[m] = { label: MESES_ABREV[m - 1], "Concluído / Não faturado": 0, "Em andamento": 0, "Faturado": 0, "Suspenso": 0 };
      doAno.forEach((c) => {
        const m = parseInt(c.dataSLA.slice(5, 7), 10);
        if (!map[m]) return;
        map[m][c.statusParcela] = (map[m][c.statusParcela] || 0) + (c.valorFaturamento || 0);
      });
      return Object.values(map);
    }
    const map = {};
    filtrados.forEach((c) => {
      if (!map[c.cliente]) map[c.cliente] = { label: c.cliente, "Concluído / Não faturado": 0, "Em andamento": 0, "Faturado": 0, "Suspenso": 0 };
      map[c.cliente][c.statusParcela] = (map[c.cliente][c.statusParcela] || 0) + (c.valorFaturamento || 0);
    });
    return Object.values(map).sort((a, b) => (b["Concluído / Não faturado"] + b["Em andamento"] + b["Faturado"] + b["Suspenso"]) - (a["Concluído / Não faturado"] + a["Em andamento"] + a["Faturado"] + a["Suspenso"])).slice(0, 10);
  }, [mes, doAno, filtrados]);

  const paginados = useMemo(() => {
    const ordenados = [...filtrados].sort((a, b) => (a.dataSLA || "").localeCompare(b.dataSLA || ""));
    return paginate(ordenados, page, pageSize);
  }, [filtrados, page, pageSize]);

  return (
    <div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 18 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6, color: COLORS.steel, fontSize: 11.5, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.04em" }}>
          <Filter size={13} /> Período
        </div>
        <Select value={ano} onChange={(v) => { setAno(v); setPage(1); }} options={anosDisponiveis} />
        <select value={mes} onChange={(e) => { setMes(e.target.value); setPage(1); }} style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "8px 10px", color: COLORS.steelLight, fontSize: 12.5, fontFamily: "'Inter', sans-serif" }}>
          <option value="Todos">Ano inteiro</option>
          {MESES_NOMES.map((nome, i) => <option key={nome} value={String(i + 1)}>{nome}</option>)}
        </select>
        {mes !== "Todos" && <span style={{ fontSize: 12.5, color: COLORS.steelLight }}>{MESES_NOMES[parseInt(mes, 10) - 1]} de {ano}</span>}
      </div>

      <div style={{ display: "flex", gap: 14, marginBottom: 20, flexWrap: "wrap" }}>
        <KpiCard icon={DollarSign} label="Total planejado" value={fmtBRL(totalPlanejado)} accent={COLORS.blue} sub={mes === "Todos" ? `ano ${ano}` : `${MESES_NOMES[parseInt(mes, 10) - 1]}/${ano}`} />
        <KpiCard icon={CheckCircle2} label="Faturado" value={fmtBRL(totalFaturado)} accent={COLORS.green} sub="parcelas já faturadas" />
        <KpiCard icon={Clock} label="Em andamento" value={fmtBRL(totalAndamento)} accent={COLORS.blue} sub="parcelas em execução" />
        <KpiCard icon={AlertTriangle} label="Concluído / Não faturado" value={fmtBRL(totalConcluidoNaoFaturado)} accent={COLORS.orange} sub="pronto, falta faturar" />
        <KpiCard icon={MinusCircle} label="Suspenso" value={fmtBRL(totalSuspenso)} accent={COLORS.steel} sub="parcelas suspensas" />
      </div>

      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 18, marginBottom: 20 }}>
        <div style={{ fontSize: 12, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600, marginBottom: 14 }}>
          {mes === "Todos" ? `Planejamento Financeiro — ano ${ano}` : `Planejamento Financeiro por cliente — ${MESES_NOMES[parseInt(mes, 10) - 1]}/${ano}`}
        </div>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={chartData} margin={{ left: 0, right: 8 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
            <XAxis dataKey="label" tick={{ fill: COLORS.steelLight, fontSize: 11 }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
            <YAxis tick={{ fill: COLORS.steel, fontSize: 10.5 }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
            <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12 }} cursor={{ fill: "rgba(255,255,255,0.03)" }} />
            <Legend wrapperStyle={{ fontSize: 11, color: COLORS.steelLight }} />
            <Bar dataKey="Concluído / Não faturado" stackId="a" fill={COLORS.orange} />
            <Bar dataKey="Em andamento" stackId="a" fill={COLORS.blue} />
            <Bar dataKey="Faturado" stackId="a" fill={COLORS.green} />
            <Bar dataKey="Suspenso" stackId="a" fill={COLORS.steel} radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr>{["Data SLA", "Cliente / Unidade", "Serviço", "Tarefa", "Técnico", "Valor", "Status parcela"].map((h) => (
              <th key={h} style={{ textAlign: "left", padding: "10px 16px", fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${COLORS.border}`, whiteSpace: "nowrap" }}>{h}</th>
            ))}</tr></thead>
            <tbody>
              {paginados.map((c) => {
                const sp = statusParcelaStyle(c.statusParcela);
                return (
                  <tr key={c.id} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                    <td style={{ padding: "10px 16px", fontSize: 12, color: COLORS.steel, fontFamily: "monospace" }}>{fmtDate(c.dataSLA)}</td>
                    <td style={{ padding: "10px 16px" }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: COLORS.ice }}>{c.cliente}</div>
                      <div style={{ fontSize: 11.5, color: COLORS.steel, marginTop: 2 }}>{c.unidade}</div>
                    </td>
                    <td style={{ padding: "10px 16px", fontSize: 12.5, color: COLORS.steelLight, maxWidth: 220 }}>{c.servico}</td>
                    <td style={{ padding: "10px 16px", fontSize: 12, color: COLORS.steel }}>{c.tarefa}</td>
                    <td style={{ padding: "10px 16px", fontSize: 12, color: COLORS.steelLight }}>{c.tecnico}</td>
                    <td style={{ padding: "10px 16px", fontSize: 12.5, color: COLORS.steelLight }}>{fmtBRL(c.valorFaturamento)}</td>
                    <td style={{ padding: "10px 16px" }}><Pill fg={sp.fg} bg={sp.bg}>{c.statusParcela}</Pill></td>
                  </tr>
                );
              })}
              {filtrados.length === 0 && <tr><td colSpan={7} style={{ padding: 40, textAlign: "center", color: COLORS.steel, fontSize: 13 }}>Nenhum contrato com Data SLA neste período.</td></tr>}
            </tbody>
          </table>
        </div>
        <Pagination page={page} setPage={setPage} pageSize={pageSize} setPageSize={setPageSize} totalItems={filtrados.length} />
      </div>
    </div>
  );
}

/* ============================================================
   PLANEJAMENTO MENSAL (Dashboard) — comparativo simples do mes:
   servicos faturados x nao faturados, por Data SLA
   ============================================================ */
function PlanejamentoMensalDashboardPage({ contratos, onOpenContrato }) {
  const hojeRef = new Date();
  const [ano, setAno] = useState(String(hojeRef.getFullYear()));
  const [mes, setMes] = useState(String(hojeRef.getMonth() + 1));
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);

  const anosDisponiveis = useMemo(() => {
    const anos = new Set(contratos.map((c) => (c.dataSLA ? c.dataSLA.slice(0, 4) : null)).filter(Boolean));
    anos.add(String(hojeRef.getFullYear()));
    return Array.from(anos).sort();
  }, [contratos]); // eslint-disable-line

  const doMes = useMemo(() => contratos.filter((c) => c.dataSLA && c.dataSLA.slice(0, 4) === ano && String(parseInt(c.dataSLA.slice(5, 7), 10)) === mes && c.statusParcela !== "Suspenso"), [contratos, ano, mes]);
  const faturados = doMes.filter((c) => c.statusParcela === "Faturado");
  const naoFaturados = doMes.filter((c) => c.statusParcela !== "Faturado");

  const valorTotal = doMes.reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const valorFaturado = faturados.reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const valorNaoFaturado = naoFaturados.reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const pctFaturado = doMes.length === 0 ? 0 : Math.round((faturados.length / doMes.length) * 100);

  const pieData = [
    { name: "Faturado", value: valorFaturado },
    { name: "Não faturado", value: valorNaoFaturado },
  ].filter((d) => d.value > 0);
  const PIE_COLORS = { "Faturado": COLORS.green, "Não faturado": COLORS.orange };

  const ordenados = useMemo(() => [...doMes].sort((a, b) => (a.dataSLA || "").localeCompare(b.dataSLA || "")), [doMes]);
  const paginados = useMemo(() => paginate(ordenados, page, pageSize), [ordenados, page, pageSize]);

  return (
    <div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 20 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6, color: COLORS.steel, fontSize: 11.5, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.04em" }}>
          <Filter size={13} /> Mês de referência
        </div>
        <Select value={ano} onChange={(v) => { setAno(v); setPage(1); }} options={anosDisponiveis} />
        <select value={mes} onChange={(e) => { setMes(e.target.value); setPage(1); }} style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "8px 10px", color: COLORS.steelLight, fontSize: 12.5, fontFamily: "'Inter', sans-serif" }}>
          {MESES_NOMES.map((nome, i) => <option key={nome} value={String(i + 1)}>{nome}</option>)}
        </select>
      </div>
      <div style={{ fontSize: 11.5, color: COLORS.steel, marginBottom: 18 }}>Serviços suspensos não entram nesta comparação — consulte-os em Contratos ou no Planejamento Financeiro.</div>

      <div style={{ display: "flex", gap: 14, marginBottom: 20, flexWrap: "wrap" }}>
        <KpiCard icon={DollarSign} label="Total do mês" value={fmtBRL(valorTotal)} accent={COLORS.blue} sub={`${doMes.length} serviço(s)`} />
        <KpiCard icon={CheckCircle2} label="Faturado" value={`${pctFaturado}%`} accent={COLORS.green} sub={`${fmtBRL(valorFaturado)} · ${faturados.length} serviço(s)`} />
        <KpiCard icon={AlertTriangle} label="Não faturado" value={fmtBRL(valorNaoFaturado)} accent={COLORS.orange} sub={`${naoFaturados.length} serviço(s)`} />
      </div>

      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 18, marginBottom: 20 }}>
        <div style={{ fontSize: 12, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600, marginBottom: 14 }}>Faturado x não faturado — {MESES_NOMES[parseInt(mes, 10) - 1]}/{ano}</div>
        <ResponsiveContainer width="100%" height={230}>
          <PieChart>
            <Pie data={pieData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={85} paddingAngle={3}>
              {pieData.map((d, i) => <Cell key={i} fill={PIE_COLORS[d.name]} />)}
            </Pie>
            <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12 }} />
            <Legend wrapperStyle={{ fontSize: 11, color: COLORS.steelLight }} />
          </PieChart>
        </ResponsiveContainer>
      </div>

      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, overflow: "hidden" }}>
        <div style={{ fontSize: 12, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600, padding: "12px 18px 10px" }}>Serviços do mês</div>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr>{["Cliente / Unidade", "Serviço", "Tarefa", "Técnico", "Valor", "Data SLA", "Status parcela"].map((h) => (
              <th key={h} style={{ textAlign: "left", padding: "10px 16px", fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${COLORS.border}` }}>{h}</th>
            ))}</tr></thead>
            <tbody>
              {paginados.map((c) => {
                const sp = statusParcelaStyle(c.statusParcela);
                return (
                  <tr key={c.id} className="row-hover" style={{ cursor: onOpenContrato ? "pointer" : "default", borderBottom: `1px solid ${COLORS.border}` }} onClick={onOpenContrato}>
                    <td style={{ padding: "10px 16px" }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: COLORS.ice }}>{c.cliente}</div>
                      <div style={{ fontSize: 11.5, color: COLORS.steel, marginTop: 2 }}>{c.unidade}</div>
                    </td>
                    <td style={{ padding: "10px 16px", fontSize: 12.5, color: COLORS.steelLight, maxWidth: 220 }}>{c.servico}</td>
                    <td style={{ padding: "10px 16px", fontSize: 12, color: COLORS.steel }}>{c.tarefa}</td>
                    <td style={{ padding: "10px 16px", fontSize: 12, color: COLORS.steelLight }}>{c.tecnico}</td>
                    <td style={{ padding: "10px 16px", fontSize: 12.5, color: COLORS.steelLight }}>{fmtBRL(c.valorFaturamento)}</td>
                    <td style={{ padding: "10px 16px", fontSize: 12, color: COLORS.steel }}>{fmtDate(c.dataSLA)}</td>
                    <td style={{ padding: "10px 16px" }}><Pill fg={sp.fg} bg={sp.bg}>{c.statusParcela}</Pill></td>
                  </tr>
                );
              })}
              {doMes.length === 0 && <tr><td colSpan={7} style={{ padding: 30, textAlign: "center", color: COLORS.steel, fontSize: 13 }}>Nenhum contrato com Data SLA neste mês.</td></tr>}
            </tbody>
          </table>
        </div>
        <Pagination page={page} setPage={setPage} pageSize={pageSize} setPageSize={setPageSize} totalItems={ordenados.length} />
      </div>
    </div>
  );
}

/* ============================================================
   CONTRATOS — listagem completa (nível de parcela/tarefa)
   ============================================================ */
function ContratosPage({ contratos, onUpdateContrato, onAddContrato, onDeleteContrato, isAdmin }) {
  const [filtroCliente, setFiltroCliente] = useState("Todos");
  const [filtroUnidade, setFiltroUnidade] = useState("Todos");
  const [filtroStatusContrato, setFiltroStatusContrato] = useState("Todos");
  const [filtroStatusParcela, setFiltroStatusParcela] = useState("Todos");
  const [busca, setBusca] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);
  const [formModal, setFormModal] = useState(null); // { title, submitLabel, initial, onSubmit } | null
  const [confirmDelete, setConfirmDelete] = useState(null); // contrato a excluir

  const clientes = useMemo(() => Array.from(new Set(contratos.map((c) => c.cliente))).sort(), [contratos]);
  const unidades = useMemo(() => {
    const base = filtroCliente === "Todos" ? contratos : contratos.filter((c) => c.cliente === filtroCliente);
    return Array.from(new Set(base.map((c) => c.unidade))).sort();
  }, [contratos, filtroCliente]);
  const statusContratoOpts = useMemo(() => Array.from(new Set(contratos.map((c) => c.statusContrato))).filter(Boolean).sort(), [contratos]);

  const filtrados = useMemo(() => contratos.filter((c) => {
    if (filtroCliente !== "Todos" && c.cliente !== filtroCliente) return false;
    if (filtroUnidade !== "Todos" && c.unidade !== filtroUnidade) return false;
    if (filtroStatusContrato !== "Todos" && c.statusContrato !== filtroStatusContrato) return false;
    if (filtroStatusParcela !== "Todos" && c.statusParcela !== filtroStatusParcela) return false;
    if (busca) {
      const q = busca.toLowerCase();
      const hay = `${c.cliente} ${c.unidade} ${c.servico} ${c.proposta}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  }), [contratos, filtroCliente, filtroUnidade, filtroStatusContrato, filtroStatusParcela, busca]);

  const totalPropostas = useMemo(() => new Set(filtrados.map((c) => c.proposta)).size, [filtrados]);
  const totalClientes = useMemo(() => new Set(filtrados.map((c) => c.cliente)).size, [filtrados]);
  const valorTotalParcelas = filtrados.reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const valorContratadoUnico = useMemo(() => {
    const seen = new Set();
    let total = 0;
    filtrados.forEach((c) => {
      const key = `${c.proposta}|${c.cliente}|${c.unidade}|${c.servico}`;
      if (seen.has(key)) return;
      seen.add(key);
      total += c.honorarios || 0;
    });
    return total;
  }, [filtrados]);
  const paginados = useMemo(() => paginate(filtrados, page, pageSize), [filtrados, page, pageSize]);

  const faturadoValor = filtrados.filter((c) => c.statusParcela === "Faturado").reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const concluidoNaoFaturadoValor = filtrados.filter((c) => c.statusParcela === "Concluído / Não faturado").reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const emAndamentoValor = filtrados.filter((c) => c.statusParcela === "Em andamento").reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const suspensoValor = filtrados.filter((c) => c.statusParcela === "Suspenso").reduce((s, c) => s + (c.valorFaturamento || 0), 0);
  const pctFaturado = valorTotalParcelas === 0 ? 0 : Math.round((faturadoValor / valorTotalParcelas) * 100);
  const pieFaturamento = [
    { name: "Faturado", value: faturadoValor },
    { name: "Concluído / Não faturado", value: concluidoNaoFaturadoValor },
    { name: "Em andamento", value: emAndamentoValor },
    { name: "Suspenso", value: suspensoValor },
  ].filter((d) => d.value > 0);

  const porServicoFaturamento = useMemo(() => {
    const map = {};
    filtrados.forEach((c) => {
      if (!map[c.servico]) map[c.servico] = { servico: c.servico, "Faturado": 0, "Concluído / Não faturado": 0, "Em andamento": 0, "Suspenso": 0 };
      map[c.servico][c.statusParcela] = (map[c.servico][c.statusParcela] || 0) + (c.valorFaturamento || 0);
    });
    return Object.values(map).sort((a, b) => (b["Faturado"] + b["Concluído / Não faturado"] + b["Em andamento"] + b["Suspenso"]) - (a["Faturado"] + a["Concluído / Não faturado"] + a["Em andamento"] + a["Suspenso"])).slice(0, 10);
  }, [filtrados]);

  const exportar = () => {
    const rows = [["Proposta", "Cliente", "Unidade", "Serviço", "Tarefa", "Tipo", "Técnico", "Coordenador", "Honorários", "Valor Faturamento", "Porcentagem", "Data SLA", "Status Serviço", "Status Parcela", "Observação"]];
    filtrados.forEach((c) => rows.push([c.proposta, c.cliente, c.unidade, c.servico, c.tarefa, c.tipo, c.tecnico, c.coordenador, c.honorarios, c.valorFaturamento, c.porcentagem, fmtDate(c.dataSLA), c.statusServico, c.statusParcela, c.observacao]));
    downloadCSV("contratos.csv", rows);
  };

  const abrirNovo = () => setFormModal({
    title: "Novo contrato", submitLabel: "Salvar contrato", initial: {},
    onSubmit: (fields) => onAddContrato(novaLinhaContrato(fields)),
  });
  const abrirAdicionarServico = (c) => setFormModal({
    title: `Adicionar serviço — contrato ${c.proposta}`, submitLabel: "Adicionar serviço",
    initial: { proposta: c.proposta, cliente: c.cliente, unidade: c.unidade, codigoLoja: c.codigoLoja, tecnico: c.tecnico, coordenador: c.coordenador, tipo: c.tipo || "Processo", servico: "", tarefa: "", honorarios: "", valorFaturamento: "", porcentagemPct: "100", dataSLA: "", statusParcela: "Concluído / Não faturado", observacao: "" },
    onSubmit: (fields) => onAddContrato(novaLinhaContrato(fields)),
  });
  const abrirEditar = (c) => setFormModal({
    title: `Editar serviço — contrato ${c.proposta}`, submitLabel: "Salvar alterações",
    initial: { proposta: c.proposta, cliente: c.cliente, unidade: c.unidade, codigoLoja: c.codigoLoja, servico: c.servico, tarefa: c.tarefa, tecnico: c.tecnico, coordenador: c.coordenador, tipo: c.tipo || "Processo", honorarios: String(c.honorarios || ""), valorFaturamento: String(c.valorFaturamento || ""), porcentagemPct: String(Math.round((c.porcentagem || 0) * 100)), dataSLA: c.dataSLA || "", statusParcela: c.statusParcela, observacao: c.observacao || "" },
    onSubmit: (fields) => onUpdateContrato(c.id, fields),
  });

  return (
    <div>
      <div style={{ display: "flex", gap: 14, marginBottom: 14, flexWrap: "wrap" }}>
        <KpiCard icon={FileSignature} label="Propostas" value={totalPropostas} accent={COLORS.blue} sub={`${filtrados.length} linha(s) de parcela`} />
        <KpiCard icon={Building2} label="Clientes" value={totalClientes} accent={COLORS.steelLight} sub="clientes no filtro atual" />
        <KpiCard icon={DollarSign} label="Valor contratado" value={fmtBRL(valorContratadoUnico)} accent={COLORS.green} sub="honorários únicos por serviço" />
      </div>
      <div style={{ display: "flex", gap: 14, marginBottom: 20, flexWrap: "wrap" }}>
        <KpiCard icon={CheckCircle2} label="Faturado" value={fmtBRL(faturadoValor)} accent={COLORS.green} sub={`${pctFaturado}% do total do filtro`} />
        <KpiCard icon={Clock} label="Concluído / Não faturado" value={fmtBRL(concluidoNaoFaturadoValor)} accent={COLORS.orange} sub="pronto, falta faturar" />
        <KpiCard icon={Timer} label="Em andamento" value={fmtBRL(emAndamentoValor)} accent={COLORS.blue} sub="ainda em execução" />
        <KpiCard icon={MinusCircle} label="Suspenso" value={fmtBRL(suspensoValor)} accent={COLORS.steel} sub="parcelas suspensas" />
      </div>

      <div style={{ display: "flex", gap: 16, marginBottom: 20, flexWrap: "wrap" }}>
        <div style={{ flex: "1 1 260px", background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 18 }}>
          <div style={{ fontSize: 12, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600, marginBottom: 14 }}>Situação das parcelas</div>
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={pieFaturamento} dataKey="value" nameKey="name" innerRadius={50} outerRadius={78} paddingAngle={3}>
                {pieFaturamento.map((d, i) => <Cell key={i} fill={statusParcelaStyle(d.name).fg} />)}
              </Pie>
              <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12 }} />
              <Legend wrapperStyle={{ fontSize: 10.5, color: COLORS.steelLight }} />
            </PieChart>
          </ResponsiveContainer>
        </div>
        <div style={{ flex: "2 1 380px", background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 18 }}>
          <div style={{ fontSize: 12, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600, marginBottom: 14 }}>Faturado, concluído, em andamento e suspenso — por serviço</div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={porServicoFaturamento} layout="vertical" margin={{ left: 8, right: 16 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} horizontal={false} />
              <XAxis type="number" tick={{ fill: COLORS.steel, fontSize: 10.5 }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
              <YAxis type="category" dataKey="servico" width={200} tick={{ fill: COLORS.steelLight, fontSize: 10.5 }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
              <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12 }} cursor={{ fill: "rgba(255,255,255,0.03)" }} />
              <Legend wrapperStyle={{ fontSize: 10.5, color: COLORS.steelLight }} />
              <Bar dataKey="Faturado" stackId="a" fill={COLORS.green} />
              <Bar dataKey="Concluído / Não faturado" stackId="a" fill={COLORS.orange} />
              <Bar dataKey="Em andamento" stackId="a" fill={COLORS.blue} />
              <Bar dataKey="Suspenso" stackId="a" fill={COLORS.steel} radius={[0, 3, 3, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "8px 12px" }}>
            <Search size={14} color={COLORS.steel} />
            <input value={busca} onChange={(e) => { setBusca(e.target.value); setPage(1); }} placeholder="Buscar cliente, unidade, serviço ou proposta..."
              style={{ background: "transparent", border: "none", outline: "none", color: COLORS.ice, fontSize: 13, width: 260 }} />
          </div>
          <Select value={filtroCliente} onChange={(v) => { setFiltroCliente(v); setFiltroUnidade("Todos"); setPage(1); }} options={clientes} placeholder="Todos os clientes" />
          <Select value={filtroUnidade} onChange={(v) => { setFiltroUnidade(v); setPage(1); }} options={unidades} placeholder="Todas as unidades" />
          <Select value={filtroStatusContrato} onChange={(v) => { setFiltroStatusContrato(v); setPage(1); }} options={statusContratoOpts} placeholder="Todo status de contrato" />
          <Select value={filtroStatusParcela} onChange={(v) => { setFiltroStatusParcela(v); setPage(1); }} options={STATUS_PARCELA_OPTIONS} placeholder="Toda status de parcela" />
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={exportar} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "9px 14px", fontSize: 13, cursor: "pointer" }}>
            <Download size={14} /> Exportar CSV
          </button>
          <button onClick={abrirNovo} style={{ display: "flex", alignItems: "center", gap: 7, background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.02em", textTransform: "uppercase" }}>
            <Plus size={15} /> Novo contrato
          </button>
        </div>
      </div>

      <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr>{["Proposta", "Cliente / Unidade", "Serviço", "Tarefa", "Tipo (editável)", "Técnico (editável)", "Honorários (editável)", "Valor parcela (editável)", "%", "Data SLA (editável)", "Status parcela", "Ações"].map((h) => (
              <th key={h} style={{ textAlign: "left", padding: "10px 16px", fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${COLORS.border}`, whiteSpace: "nowrap" }}>{h}</th>
            ))}</tr></thead>
            <tbody>
              {paginados.map((c) => {
                const sp = statusParcelaStyle(c.statusParcela);
                return (
                  <tr key={c.id} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                    <td style={{ padding: "10px 16px", fontSize: 12, color: COLORS.steel, fontFamily: "monospace" }}>{c.proposta}</td>
                    <td style={{ padding: "10px 16px" }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: COLORS.ice }}>{c.cliente}</div>
                      <div style={{ fontSize: 11.5, color: COLORS.steel, marginTop: 2 }}>{c.unidade}</div>
                    </td>
                    <td style={{ padding: "10px 16px", fontSize: 12.5, color: COLORS.steelLight, maxWidth: 220 }}>{c.servico}</td>
                    <td style={{ padding: "10px 16px", fontSize: 12, color: COLORS.steel }}>{c.tarefa}</td>
                    <td style={{ padding: "10px 16px" }}>
                      <select value={c.tipo || "Processo"} onChange={(e) => onUpdateContrato(c.id, { tipo: e.target.value })}
                        style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "5px 7px", color: COLORS.ice, fontSize: 12 }}>
                        {CONTRATO_TIPO_OPTIONS.map((o) => <option key={o} value={o}>{o}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: "10px 16px" }}>
                      <select value={TECNICOS_OPTIONS.includes(c.tecnico) ? c.tecnico : ""} onChange={(e) => onUpdateContrato(c.id, { tecnico: e.target.value })}
                        style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "5px 7px", color: COLORS.ice, fontSize: 12, width: 130 }}>
                        <option value="">Selecionar...</option>
                        {TECNICOS_OPTIONS.map((o) => <option key={o} value={o}>{o}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: "10px 16px" }}>
                      {isAdmin ? (
                        <input type="number" value={c.honorarios} onChange={(e) => onUpdateContrato(c.id, { honorarios: parseFloat(e.target.value) || 0 })}
                          style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "5px 7px", color: COLORS.ice, fontSize: 12, width: 100 }} />
                      ) : <span style={{ fontSize: 12, color: COLORS.steel }}>••••••</span>}
                    </td>
                    <td style={{ padding: "10px 16px" }}>
                      {isAdmin ? (
                        <input type="number" value={c.valorFaturamento} onChange={(e) => onUpdateContrato(c.id, { valorFaturamento: parseFloat(e.target.value) || 0 })}
                          style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "5px 7px", color: COLORS.ice, fontSize: 12, width: 100 }} />
                      ) : <span style={{ fontSize: 12, color: COLORS.steel }}>••••••</span>}
                    </td>
                    <td style={{ padding: "10px 16px", fontSize: 12, color: COLORS.steel }}>{Math.round((c.porcentagem || 0) * 100)}%</td>
                    <td style={{ padding: "10px 16px" }}>
                      <input type="date" value={c.dataSLA || ""} onChange={(e) => onUpdateContrato(c.id, { dataSLA: e.target.value })}
                        style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "5px 7px", color: COLORS.ice, fontSize: 12 }} />
                    </td>
                    <td style={{ padding: "10px 16px" }}>
                      <select value={c.statusParcela} onChange={(e) => onUpdateContrato(c.id, { statusParcela: e.target.value })}
                        style={{ background: sp.bg, color: sp.fg, border: `1px solid ${sp.fg}55`, borderRadius: 999, padding: "4px 9px", fontSize: 11, fontWeight: 700 }}>
                        {STATUS_PARCELA_OPTIONS.map((o) => <option key={o} value={o} style={{ background: COLORS.panel, color: COLORS.ice }}>{o}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: "10px 16px" }}>
                      <div style={{ display: "flex", gap: 6 }}>
                        <button onClick={() => abrirAdicionarServico(c)} title="Adicionar serviço a este contrato"
                          style={{ background: "transparent", border: `1px solid ${COLORS.border}`, borderRadius: 6, width: 26, height: 26, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" }}>
                          <PlusCircle size={13} color={COLORS.steelLight} />
                        </button>
                        <button onClick={() => abrirEditar(c)} title="Editar este serviço"
                          style={{ background: "transparent", border: `1px solid ${COLORS.border}`, borderRadius: 6, width: 26, height: 26, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" }}>
                          <Pencil size={12} color={COLORS.steelLight} />
                        </button>
                        <button onClick={() => setConfirmDelete(c)} title="Excluir este serviço do contrato"
                          style={{ background: "transparent", border: `1px solid ${COLORS.red}55`, borderRadius: 6, width: 26, height: 26, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" }}>
                          <Trash2 size={12} color={COLORS.red} />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
              {filtrados.length === 0 && (
                <tr><td colSpan={12} style={{ padding: 40, textAlign: "center", color: COLORS.steel, fontSize: 13 }}>
                  Nenhum contrato importado ainda. Use "Importar novos clientes/contratos" no menu Clientes.
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
        <Pagination page={page} setPage={setPage} pageSize={pageSize} setPageSize={setPageSize} totalItems={filtrados.length} />
      </div>
      {formModal && <ContratoFormModal title={formModal.title} submitLabel={formModal.submitLabel} initial={formModal.initial} onSubmit={formModal.onSubmit} onClose={() => setFormModal(null)} clientesExistentes={clientes} isAdmin={isAdmin} />}
      {confirmDelete && (
        <ModalShell title="Excluir serviço do contrato" onClose={() => setConfirmDelete(null)} maxWidth={420}>
          <p style={{ fontSize: 13, color: COLORS.steelLight, lineHeight: 1.6 }}>
            Remover <b style={{ color: COLORS.ice }}>{confirmDelete.servico}</b> ({confirmDelete.tarefa}) do contrato <b style={{ color: COLORS.ice }}>{confirmDelete.proposta}</b> — {confirmDelete.cliente}? Esta ação não pode ser desfeita.
          </p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
            <button onClick={() => setConfirmDelete(null)} style={{ background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.steelLight, borderRadius: 7, padding: "9px 16px", fontSize: 13, cursor: "pointer" }}>Cancelar</button>
            <button onClick={() => { onDeleteContrato(confirmDelete.id); setConfirmDelete(null); }}
              style={{ background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.03em", textTransform: "uppercase" }}>
              Excluir
            </button>
          </div>
        </ModalShell>
      )}
    </div>
  );
}

/* ============================================================
   APP
   ============================================================ */
function ControleProcessos({ usuarioLogado, onLogout }) {
  const isAdmin = usuarioLogado.role === "admin";
  OCULTAR_VALORES = !isAdmin;
  const [processos, setProcessos] = useState(MOCK_PROCESSOS);
  const [tab, setTab] = useState("dashboard-processos");
  const [dashboardOpen, setDashboardOpen] = useState(true);
  const [clientesOpen, setClientesOpen] = useState(true);
  const [busca, setBusca] = useState("");
  const [filtros, setFiltros] = useState({ cliente: "Todos", unidade: "Todos", assunto: "Todos", responsavel: "Todos" });
  const [selected, setSelected] = useState(null);
  const [showNew, setShowNew] = useState(false);
  const [contratos, setContratos] = useState([]);
  const [agendaItens, setAgendaItens] = useState([]);
  const [carregandoDados, setCarregandoDados] = useState(true);

  // Carrega tudo do banco de dados assim que a tela abre
  useEffect(() => {
    let ativo = true;
    async function carregar() {
      const [rp, rc, ra] = await Promise.all([
        supabase.from("processos").select("*").order("created_at", { ascending: false }),
        supabase.from("contratos").select("*").order("created_at", { ascending: false }),
        supabase.from("agenda_itens").select("*").order("data"),
      ]);
      if (!ativo) return;
      if (!rp.error && rp.data) setProcessos(rp.data.map(rowToProcesso));
      if (!rc.error && rc.data) setContratos(rc.data.map(rowToContrato));
      if (!ra.error && ra.data) setAgendaItens(ra.data.map(rowToAgendaItem));
      setCarregandoDados(false);
    }
    carregar();
    return () => { ativo = false; };
  }, []);

  const addAgendaItem = async (item) => {
    const { id, ...campos } = item;
    const { data, error } = await supabase.from("agenda_itens").insert(campos).select().single();
    if (!error && data) setAgendaItens((prev) => [...prev, rowToAgendaItem(data)]);
  };
  const removeAgendaItem = async (id) => {
    setAgendaItens((prev) => prev.filter((a) => a.id !== id));
    await supabase.from("agenda_itens").delete().eq("id", id);
  };
  const updateContrato = (id, fields) => {
    setContratos((prev) => {
      const next = prev.map((c) => (c.id === id ? { ...c, ...fields } : c));
      if (fields.tecnico !== undefined || fields.tipo !== undefined) {
        const atualizado = next.find((c) => c.id === id);
        if (atualizado) setProcessos((pprev) => {
          const sincronizados = sincronizarTecnicoTipoProcesso(atualizado, pprev);
          sincronizados.forEach((p, i) => { if (p !== pprev[i]) supabase.from("processos").update({ tecnico: p.tecnico, tipo: p.tipo }).eq("id", p.id).then(() => {}); });
          return sincronizados;
        });
      }
      return next;
    });
    supabase.from("contratos").update(contratoToRow({ id, ...fields })).eq("id", id).then(() => {});
  };
  const deleteContrato = (id) => {
    setContratos((prev) => prev.filter((c) => c.id !== id));
    supabase.from("contratos").delete().eq("id", id).then(() => {});
  };
  const addContratoManual = async (row, vincularProcesso) => {
    const { id, ...campos } = row;
    const { data, error } = await supabase.from("contratos").insert(contratoToRow(campos)).select().single();
    const salvo = (!error && data) ? rowToContrato(data) : row;
    setContratos((prev) => [salvo, ...prev]);
    if (vincularProcesso) await persistirMergeProcessos([salvo]);
  };

  // Aplica o merge de contrato(s) → processo(s) tanto no estado local quanto no banco
  const persistirMergeProcessos = async (linhasContrato) => {
    let atualizadoLocal;
    setProcessos((prev) => { atualizadoLocal = mergeContratosIntoProcessos(prev, linhasContrato); return atualizadoLocal; });
    for (const p of atualizadoLocal || []) {
      const jaExistiaAntes = processos.some((x) => x.id === p.id);
      if (!jaExistiaAntes) {
        const { id, ...campos } = processoToRow(p);
        const { data } = await supabase.from("processos").insert(campos).select().single();
        if (data) setProcessos((prev) => prev.map((x) => (x.id === p.id ? rowToProcesso(data) : x)));
      } else {
        await supabase.from("processos").update(processoToRow(p)).eq("id", p.id);
      }
    }
  };

  const [processosPage, setProcessosPage] = useState(1);
  const [processosPageSize, setProcessosPageSize] = useState(10);
  useEffect(() => { setProcessosPage(1); }, [busca, filtros]);
  const isDashboard = tab === "dashboard-processos" || tab === "dashboard-financeiro" || tab === "dashboard-planejamento";
  const isClientes = tab === "clientes" || tab === "unidades" || tab === "contratos" || tab === "importar-contratos";
  const updateProcesso = (novo) => {
    setProcessos((prev) => prev.map((p) => (p.id === novo.id ? novo : p)));
    supabase.from("processos").update(processoToRow(novo)).eq("id", novo.id).then(() => {});
  };
  const concluirProcesso = (novo) => {
    updateProcesso(novo);
    setSelected(novo);
    setContratos((prev) => {
      const sincronizados = sincronizarParcelasNaConclusao(novo, prev);
      sincronizados.forEach((c, i) => { if (c !== prev[i]) supabase.from("contratos").update({ status_parcela: c.statusParcela }).eq("id", c.id).then(() => {}); });
      return sincronizados;
    });
  };

  // Concilia uma planilha importada com o que já está salvo no banco:
  // atualiza as linhas que já existem, cria as que são novas.
  const importarContratosPersistindo = async (rows) => {
    const chave = (c) => `${c.proposta}|${c.cliente}|${c.unidade}|${c.servico}|${c.tarefa}`;
    const existentesPorChave = {};
    contratos.forEach((c) => { existentesPorChave[chave(c)] = c; });

    const salvos = [];
    for (const nova of rows) {
      const existente = existentesPorChave[chave(nova)];
      const { id, ...campos } = contratoToRow(nova);
      if (existente) {
        const { data } = await supabase.from("contratos").update(campos).eq("id", existente.id).select().single();
        salvos.push(data ? rowToContrato(data) : { ...existente, ...nova, id: existente.id });
      } else {
        const { data } = await supabase.from("contratos").insert(campos).select().single();
        salvos.push(data ? rowToContrato(data) : nova);
      }
    }
    setContratos((prev) => {
      const chavesNovas = new Set(salvos.map(chave));
      return [...prev.filter((c) => !chavesNovas.has(chave(c))), ...salvos];
    });
    await persistirMergeProcessos(salvos);
  };

  const filtrados = useMemo(() => applyFiltros(processos, filtros), [processos, filtros]);
  const buscados = useMemo(() => {
    if (!busca) return filtrados;
    const q = busca.toLowerCase();
    return filtrados.filter((p) => `${p.cliente} ${p.unidade} ${p.assunto} ${p.numero}`.toLowerCase().includes(q));
  }, [filtrados, busca]);

  // KPIs (sobre o conjunto filtrado)
  const total = filtrados.length;
  const comPrimers = filtrados.filter((p) => STATUS_CONFIG[p.statusAtual].responsavel === "Primers").length;
  const comCliente = filtrados.filter((p) => STATUS_CONFIG[p.statusAtual].responsavel === "Cliente").length;
  const comOrgao = filtrados.filter((p) => STATUS_CONFIG[p.statusAtual].responsavel === "Órgão").length;
  const vencidos = filtrados.filter((p) => { const dr = diasRestantes(p); return dr !== null && dr < 0 && !STATUS_CONFIG[p.statusAtual].final; }).length;

  const administrativos = filtrados.filter((p) => p.tipo === "Processo").length;
  const tecnicos = filtrados.filter((p) => p.tipo === "Serviço Técnico").length;
  const emAnalise = filtrados.filter((p) => STATUS_CONFIG[p.statusAtual].grupo === "analise").length;
  const emExigencia = filtrados.filter((p) => STATUS_CONFIG[p.statusAtual].grupo === "exigencia").length;
  const aguardandoInicio = filtrados.filter((p) => STATUS_CONFIG[p.statusAtual].grupo === "aguardando").length;
  const indeferidos = filtrados.filter((p) => STATUS_CONFIG[p.statusAtual].grupo === "indeferido").length;

  const bloqueados = useMemo(() => filtrados.map((p) => ({ p, bloqueio: processoBloqueado(p, processos) })).filter((x) => x.bloqueio), [filtrados, processos]);

  // Gráfico 1: por cliente OU por unidade (se um cliente estiver selecionado)
  const porClienteOuUnidade = useMemo(() => {
    const agruparPor = filtros.cliente !== "Todos" ? "unidade" : "cliente";
    const map = {};
    filtrados.forEach((p) => { map[p[agruparPor]] = (map[p[agruparPor]] || 0) + 1; });
    return { label: agruparPor === "unidade" ? "Processos por unidade" : "Processos por cliente", data: Object.entries(map).map(([k, qtd]) => ({ k, qtd })).sort((a, b) => b.qtd - a.qtd) };
  }, [filtrados, filtros.cliente]);

  // Gráfico 2: por responsável
  const porResponsavel = useMemo(() => {
    return RESPONSAVEIS.map((r) => ({ name: r, value: filtrados.filter((p) => STATUS_CONFIG[p.statusAtual].responsavel === r).length })).filter((r) => r.value > 0);
  }, [filtrados]);
  const RESP_COLOR = { Primers: COLORS.orange, Cliente: COLORS.yellow, Órgão: COLORS.blue, Finalizado: COLORS.green };

  // Gráfico 3: por tipo de serviço (assunto)
  const porAssunto = useMemo(() => {
    const map = {};
    filtrados.forEach((p) => { map[p.assunto] = (map[p.assunto] || 0) + 1; });
    return Object.entries(map).map(([assunto, qtd]) => ({ assunto, qtd })).sort((a, b) => b.qtd - a.qtd).slice(0, 8);
  }, [filtrados]);

  const urgentes = useMemo(() => {
    return filtrados.filter((p) => !STATUS_CONFIG[p.statusAtual].final)
      .map((p) => ({ ...p, dr: diasRestantes(p) })).filter((p) => p.dr !== null)
      .sort((a, b) => a.dr - b.dr).slice(0, 6);
  }, [filtrados]);

  return (
    <div style={{ fontFamily: "'Inter', sans-serif", background: COLORS.bg, minHeight: "100%", color: COLORS.ice, display: "flex", flexDirection: "column", position: "relative" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        ::placeholder { color: ${COLORS.steel}; opacity: 0.7; }
        .blueprint-bg { background-image: linear-gradient(${COLORS.border} 1px, transparent 1px), linear-gradient(90deg, ${COLORS.border} 1px, transparent 1px); background-size: 42px 42px; }
        .nav-item { transition: background .15s, color .15s; cursor: pointer; }
        .row-hover:hover { background: ${COLORS.panelSoft} !important; }
        table { border-collapse: collapse; width: 100%; }
        th { position: sticky; top: 0; background: ${COLORS.panelAlt}; z-index: 1; }
        select option { background: ${COLORS.panel}; }
      `}</style>

      <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
      {/* SIDEBAR */}
      <aside style={{ width: 230, background: COLORS.panel, borderRight: `1px solid ${COLORS.border}`, padding: "22px 16px", flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 3, marginBottom: 4 }}>
          <div style={{ width: 5, height: 22, background: "#3a4a63" }} />
          <div style={{ width: 5, height: 22, background: "#93a2b5" }} />
          <div style={{ width: 5, height: 22, background: "#d7dee6" }} />
          <div style={{ width: 5, height: 22, background: COLORS.red }} />
          <span style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 700, fontSize: 19, color: COLORS.red, marginLeft: 8, letterSpacing: "0.02em" }}>PRIMERS</span>
        </div>
        <div style={{ fontSize: 10.5, color: COLORS.steel, letterSpacing: "0.06em", marginBottom: 22, marginLeft: 2 }}>CONTROL</div>
        <nav style={{ display: "flex", flexDirection: "column", gap: 3 }}>
          <div className="nav-item" onClick={() => setClientesOpen((o) => !o)} style={{
            display: "flex", alignItems: "center", justifyContent: "space-between", padding: "9px 10px", borderRadius: 7,
            background: isClientes ? COLORS.redDim : "transparent", color: isClientes ? COLORS.red : COLORS.steelLight,
            fontSize: 13.5, fontWeight: isClientes ? 700 : 500,
          }}>
            <span style={{ display: "flex", alignItems: "center", gap: 10 }}><Building2 size={16} />Clientes</span>
            {clientesOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
          </div>
          {clientesOpen && (
            <div style={{ display: "flex", flexDirection: "column", gap: 2, marginLeft: 14, marginBottom: 4 }}>
              {[
                { id: "clientes", label: "Clientes", icon: Building2 },
                { id: "unidades", label: "Unidades de clientes", icon: FileStack },
                { id: "contratos", label: "Contratos", icon: FileSignature },
                { id: "importar-contratos", label: "Importar novos clientes/contratos", icon: Upload },
              ].map((item) => (
                <div key={item.id} className="nav-item" onClick={() => setTab(item.id)} style={{
                  display: "flex", alignItems: "center", gap: 9, padding: "7px 10px", borderRadius: 7,
                  background: tab === item.id ? COLORS.redDim : "transparent", color: tab === item.id ? COLORS.red : COLORS.steel,
                  fontSize: 12.5, fontWeight: tab === item.id ? 700 : 500, borderLeft: `2px solid ${tab === item.id ? COLORS.red : COLORS.border}`,
                }}><item.icon size={13} />{item.label}</div>
              ))}
            </div>
          )}

          <div className="nav-item" onClick={() => setDashboardOpen((o) => !o)} style={{
            display: "flex", alignItems: "center", justifyContent: "space-between", padding: "9px 10px", borderRadius: 7,
            background: isDashboard ? COLORS.redDim : "transparent", color: isDashboard ? COLORS.red : COLORS.steelLight,
            fontSize: 13.5, fontWeight: isDashboard ? 700 : 500,
          }}>
            <span style={{ display: "flex", alignItems: "center", gap: 10 }}><LayoutDashboard size={16} />Dashboard</span>
            {dashboardOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
          </div>
          {dashboardOpen && (
            <div style={{ display: "flex", flexDirection: "column", gap: 2, marginLeft: 14, marginBottom: 4 }}>
              {[
                { id: "dashboard-processos", label: "Processos", icon: FileStack },
                { id: "dashboard-financeiro", label: "Planejamento Mensal", icon: DollarSign },
                ...(isAdmin ? [{ id: "dashboard-planejamento", label: "Planejamento Financeiro", icon: Clock }] : []),
              ].map((item) => (
                <div key={item.id} className="nav-item" onClick={() => setTab(item.id)} style={{
                  display: "flex", alignItems: "center", gap: 9, padding: "7px 10px", borderRadius: 7,
                  background: tab === item.id ? COLORS.redDim : "transparent", color: tab === item.id ? COLORS.red : COLORS.steel,
                  fontSize: 12.5, fontWeight: tab === item.id ? 700 : 500, borderLeft: `2px solid ${tab === item.id ? COLORS.red : COLORS.border}`,
                }}><item.icon size={13} />{item.label}</div>
              ))}
            </div>
          )}
          {[
            { id: "processos", label: "Controle de Processos", icon: ListChecks },
            { id: "atualizacoes", label: "Status de Serviço", icon: History },
            ...(isAdmin ? [{ id: "acessos", label: "Gerenciar acessos", icon: Users }] : []),
          ].map((item) => (
            <div key={item.id} className="nav-item" onClick={() => setTab(item.id)} style={{
              display: "flex", alignItems: "center", gap: 10, padding: "9px 10px", borderRadius: 7,
              background: tab === item.id ? COLORS.redDim : "transparent", color: tab === item.id ? COLORS.red : COLORS.steelLight,
              fontSize: 13.5, fontWeight: tab === item.id ? 700 : 500,
            }}><item.icon size={16} />{item.label}</div>
          ))}
        </nav>
        <div style={{ marginTop: 32, padding: "12px 12px", background: COLORS.panelAlt, borderRadius: 8, border: `1px solid ${COLORS.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
          <div>
            <div style={{ fontSize: 12.5, color: COLORS.ice, fontWeight: 700 }}>{usuarioLogado.nome}</div>
            <div style={{ fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em" }}>{isAdmin ? "Administrador" : "Operacional"}</div>
          </div>
          <button onClick={onLogout} title="Sair" style={{ background: "transparent", border: `1px solid ${COLORS.border}`, borderRadius: 6, width: 28, height: 28, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 }}>
            <LogOut size={13} color={COLORS.steel} />
          </button>
        </div>
      </aside>

      {/* MAIN */}
      <main className="blueprint-bg" style={{ flex: 1, minWidth: 0, padding: "22px 28px 40px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 12 }}>
          <div>
            <h1 style={{ fontFamily: "'Oswald', sans-serif", fontSize: 22, fontWeight: 600, color: COLORS.ice, textTransform: "uppercase", letterSpacing: "0.02em" }}>
              {tab === "dashboard-processos" && "Dashboard · Processos"}
              {tab === "dashboard-financeiro" && "Dashboard · Planejamento Mensal"}
              {tab === "dashboard-planejamento" && "Dashboard · Planejamento Financeiro"}
              {tab === "clientes" && "Clientes"}
              {tab === "unidades" && "Unidades de clientes"}
              {tab === "contratos" && "Contratos"}
              {tab === "importar-contratos" && "Importar novos clientes/contratos"}
              {tab === "processos" && "Controle de Processos"}
              {tab === "atualizacoes" && "Status de Serviço"}
              {tab === "acessos" && "Gerenciar acessos"}
            </h1>
            <p style={{ fontSize: 12.5, color: COLORS.steel, marginTop: 2 }}>
              {tab === "dashboard-processos" && `${total} de ${processos.length} processos totais no filtro atual`}
              {tab === "dashboard-financeiro" && "Serviços faturados x não faturados, mês a mês"}
              {tab === "dashboard-planejamento" && "Financeiro previsto e realizado por mês, conforme a Data SLA de cada contrato"}
              {tab === "clientes" && "Clientes reconhecidos a partir dos contratos importados"}
              {tab === "unidades" && "Unidades reconhecidas a partir dos contratos importados"}
              {tab === "contratos" && `${contratos.length} linha(s) de contrato importada(s)`}
              {tab === "importar-contratos" && "Envie a planilha para atualizar clientes, unidades e contratos"}
              {tab === "processos" && `${buscados.length} de ${processos.length} processo(s)`}
              {tab === "atualizacoes" && "Situação de cada serviço, com base no que foi registrado em Controle de Processos"}
              {tab === "acessos" && "Criar ou remover logins de administradores e operacionais"}
            </p>
          </div>
          {(tab === "processos" || tab === "dashboard-processos") && (
            <div style={{ display: "flex", gap: 10 }}>
              {tab === "processos" && (
                <button onClick={() => setShowNew(true)} style={{ display: "flex", alignItems: "center", gap: 7, background: COLORS.red, border: "none", color: "#fff", borderRadius: 7, padding: "9px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'Oswald', sans-serif", letterSpacing: "0.02em", textTransform: "uppercase" }}>
                  <Plus size={15} /> Novo processo
                </button>
              )}
            </div>
          )}
        </div>

        {(tab === "dashboard-processos" || tab === "processos") && <FilterBar processos={processos} filtros={filtros} setFiltros={setFiltros} />}

        {tab === "dashboard-processos" && (
          <>
            <div style={{ display: "flex", gap: 14, marginBottom: 14, flexWrap: "wrap" }}>
              <KpiCard icon={FileStack} label="Total (filtro atual)" value={total} accent={COLORS.blue} sub={`${processos.length} no total geral`} />
              <KpiCard icon={Clock} label="Com a Primers" value={comPrimers} accent={COLORS.orange} sub="nossa responsabilidade" />
              <KpiCard icon={Timer} label="Com o cliente" value={comCliente} accent={COLORS.yellow} sub="aguardando retorno" />
              <KpiCard icon={Building2} label="Com o órgão" value={comOrgao} accent={COLORS.blue} sub="aguardando análise" />
              <KpiCard icon={AlertTriangle} label="Vencidos" value={vencidos} accent={COLORS.red} sub="prazo já expirado" />
            </div>
            <div style={{ display: "flex", gap: 14, marginBottom: 20, flexWrap: "wrap" }}>
              <KpiCard icon={ListChecks} label="Processo" value={administrativos} accent={COLORS.steelLight} sub="processos administrativos" />
              <KpiCard icon={Wrench} label="Serviço Técnico" value={tecnicos} accent={COLORS.steelLight} sub="serviços técnicos" />
              <KpiCard icon={Clock} label="Aguardando início" value={aguardandoInicio} accent={COLORS.steel} sub="importados, ainda não iniciados" />
              <KpiCard icon={Search} label="Em análise" value={emAnalise} accent={COLORS.blue} sub="protocolado / aguardando órgão" />
              <KpiCard icon={AlertTriangle} label="Em exigência" value={emExigencia} accent={COLORS.orange} sub="Primers ou cliente" />
              <KpiCard icon={XCircle} label="Indeferidos" value={indeferidos} accent={COLORS.overdue} sub="negados pelo órgão" />
            </div>

            <AgendaSemanal processos={filtrados} agendaItens={agendaItens} onOpenProcesso={(p) => setSelected(p)} onAddItem={addAgendaItem} onRemoveItem={removeAgendaItem} />

            <div style={{ display: "flex", gap: 16, marginBottom: 20, flexWrap: "wrap" }}>
              <div style={{ flex: "2 1 380px", background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 18 }}>
                <div style={{ fontSize: 12, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600, marginBottom: 14 }}>{porClienteOuUnidade.label}</div>
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={porClienteOuUnidade.data} layout="vertical" margin={{ left: 8, right: 16 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} horizontal={false} />
                    <XAxis type="number" tick={{ fill: COLORS.steel, fontSize: 11 }} axisLine={{ stroke: COLORS.border }} tickLine={false} allowDecimals={false} />
                    <YAxis type="category" dataKey="k" width={140} tick={{ fill: COLORS.steelLight, fontSize: 11 }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
                    <Tooltip contentStyle={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12 }} labelStyle={{ color: COLORS.ice }} cursor={{ fill: "rgba(255,255,255,0.03)" }} />
                    <Bar dataKey="qtd" fill={COLORS.red} radius={[0, 3, 3, 0]} barSize={14} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div style={{ flex: "1 1 260px", background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 18 }}>
                <div style={{ fontSize: 12, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600, marginBottom: 14 }}>Por responsabilidade</div>
                <ResponsiveContainer width="100%" height={240}>
                  <PieChart>
                    <Pie data={porResponsavel} dataKey="value" nameKey="name" innerRadius={52} outerRadius={80} paddingAngle={3}>
                      {porResponsavel.map((entry, i) => <Cell key={i} fill={RESP_COLOR[entry.name]} />)}
                    </Pie>
                    <Tooltip contentStyle={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12 }} />
                    <Legend wrapperStyle={{ fontSize: 11, color: COLORS.steelLight }} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 18, marginBottom: 20 }}>
              <div style={{ fontSize: 12, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600, marginBottom: 14 }}>Por tipo de serviço</div>
              <ResponsiveContainer width="100%" height={Math.max(160, porAssunto.length * 34)}>
                <BarChart data={porAssunto} layout="vertical" margin={{ left: 8, right: 16 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} horizontal={false} />
                  <XAxis type="number" tick={{ fill: COLORS.steel, fontSize: 11 }} axisLine={{ stroke: COLORS.border }} tickLine={false} allowDecimals={false} />
                  <YAxis type="category" dataKey="assunto" width={260} tick={{ fill: COLORS.steelLight, fontSize: 11 }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
                  <Tooltip contentStyle={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12 }} cursor={{ fill: "rgba(255,255,255,0.03)" }} />
                  <Bar dataKey="qtd" fill={COLORS.blue} radius={[0, 3, 3, 0]} barSize={14} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            {bloqueados.length > 0 && (
              <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: "6px 0 4px", marginBottom: 20 }}>
                <div style={{ fontSize: 12, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600, padding: "12px 18px 10px", display: "flex", alignItems: "center", gap: 7 }}>
                  <Lock size={13} /> Bloqueados por dependência de outro processo
                </div>
                <table>
                  <thead><tr>{["Processo", "Cliente", "Depende da conclusão de", "Status da dependência"].map((h) => <th key={h} style={{ textAlign: "left", padding: "8px 18px", fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${COLORS.border}` }}>{h}</th>)}</tr></thead>
                  <tbody>
                    {bloqueados.map(({ p, bloqueio }) => {
                      const stB = STATUS_CONFIG[bloqueio.statusAtual];
                      return (
                        <tr key={p.id} className="row-hover" style={{ cursor: "pointer" }} onClick={() => setSelected(p)}>
                          <td style={{ padding: "10px 18px", fontSize: 13, color: COLORS.steelLight, borderBottom: `1px solid ${COLORS.border}` }}>{p.assunto}</td>
                          <td style={{ padding: "10px 18px", fontSize: 13, borderBottom: `1px solid ${COLORS.border}` }}>{p.cliente}</td>
                          <td style={{ padding: "10px 18px", fontSize: 13, color: COLORS.steelLight, borderBottom: `1px solid ${COLORS.border}` }}>{bloqueio.assunto}</td>
                          <td style={{ padding: "10px 18px", borderBottom: `1px solid ${COLORS.border}` }}><Pill fg={stB.fg} bg={stB.bg}>{stB.label}</Pill></td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: "6px 0 4px" }}>
              <div style={{ fontSize: 12, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600, padding: "12px 18px 10px" }}>Mais urgentes</div>
              <table>
                <thead><tr>{["Cliente", "Assunto", "Status", "Prazo", "Técnico"].map((h) => <th key={h} style={{ textAlign: "left", padding: "8px 18px", fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${COLORS.border}` }}>{h}</th>)}</tr></thead>
                <tbody>
                  {urgentes.map((p) => {
                    const st = STATUS_CONFIG[p.statusAtual]; const prazo = prazoInfo(p.dr);
                    return (
                      <tr key={p.id} className="row-hover" style={{ cursor: "pointer" }} onClick={() => setSelected(p)}>
                        <td style={{ padding: "10px 18px", fontSize: 13, borderBottom: `1px solid ${COLORS.border}` }}>{p.cliente}</td>
                        <td style={{ padding: "10px 18px", fontSize: 13, color: COLORS.steelLight, borderBottom: `1px solid ${COLORS.border}` }}>{p.assunto}</td>
                        <td style={{ padding: "10px 18px", borderBottom: `1px solid ${COLORS.border}` }}><Pill fg={st.fg} bg={st.bg}>{st.label}</Pill></td>
                        <td style={{ padding: "10px 18px", borderBottom: `1px solid ${COLORS.border}` }}><Pill fg={prazo.fg} bg={prazo.bg}>{prazo.label}</Pill></td>
                        <td style={{ padding: "10px 18px", fontSize: 13, color: COLORS.steel, borderBottom: `1px solid ${COLORS.border}` }}>{p.tecnico || "—"}</td>
                      </tr>
                    );
                  })}
                  {urgentes.length === 0 && <tr><td colSpan={5} style={{ padding: 20, textAlign: "center", color: COLORS.steel, fontSize: 13 }}>Nenhum processo com prazo em aberto.</td></tr>}
                </tbody>
              </table>
            </div>
          </>
        )}

        {tab === "dashboard-financeiro" && <PlanejamentoMensalDashboardPage contratos={contratos} onOpenContrato={() => setTab("contratos")} />}

        {tab === "processos" && (
          <>
            <div style={{ display: "flex", alignItems: "center", gap: 8, background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "8px 12px", marginBottom: 16, maxWidth: 420 }}>
              <Search size={14} color={COLORS.steel} />
              <input value={busca} onChange={(e) => setBusca(e.target.value)} placeholder="Buscar por cliente, assunto ou nº do processo..."
                style={{ background: "transparent", border: "none", outline: "none", color: COLORS.ice, fontSize: 13, width: "100%" }} />
            </div>
            <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10, overflow: "hidden" }}>
              <div style={{ overflowX: "auto" }}>
                <table>
                  <thead><tr>{["Cliente / Unidade", "Assunto", "Tipo", "Técnico", "Nº processo", "Status", "Protocolo", "Previsão órgão", "Prazo", "Atualização", ""].map((h) => <th key={h} style={{ textAlign: "left", padding: "10px 16px", fontSize: 10.5, color: COLORS.steel, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${COLORS.border}`, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
                  <tbody>
                    {paginate(buscados, processosPage, processosPageSize).map((p) => {
                      const st = STATUS_CONFIG[p.statusAtual]; const dr = diasRestantes(p); const prazo = prazoInfo(dr); const ds = diasSemAtualizacao(p);
                      const bloqueio = processoBloqueado(p, processos);
                      const aguardandoInicioRow = p.statusAtual === "aguardando";
                      return (
                        <tr key={p.id} className="row-hover" style={{ cursor: "pointer" }} onClick={() => setSelected(p)}>
                          <td style={{ padding: "11px 16px", borderBottom: `1px solid ${COLORS.border}` }}>
                            <div style={{ fontSize: 13, fontWeight: 600, color: COLORS.ice, display: "flex", alignItems: "center", gap: 6 }}><Building2 size={12} color={COLORS.steel} />{p.cliente}</div>
                            <div style={{ fontSize: 11.5, color: COLORS.steel, marginTop: 2 }}>{p.unidade}</div>
                          </td>
                          <td style={{ padding: "11px 16px", fontSize: 12.5, color: COLORS.steelLight, borderBottom: `1px solid ${COLORS.border}`, maxWidth: 220 }}>
                            {p.assunto}{bloqueio && <Lock size={11} color={COLORS.red} style={{ marginLeft: 6, verticalAlign: "middle" }} />}
                          </td>
                          <td style={{ padding: "11px 16px", fontSize: 11.5, color: COLORS.steel, borderBottom: `1px solid ${COLORS.border}`, whiteSpace: "nowrap" }}>{p.tipo}</td>
                          <td style={{ padding: "11px 16px", fontSize: 11.5, color: COLORS.steelLight, borderBottom: `1px solid ${COLORS.border}`, whiteSpace: "nowrap" }}>{p.tecnico}</td>
                          <td style={{ padding: "11px 16px", fontSize: 12, color: COLORS.steel, borderBottom: `1px solid ${COLORS.border}`, fontFamily: "monospace" }}>{p.numero}</td>
                          <td style={{ padding: "11px 16px", borderBottom: `1px solid ${COLORS.border}` }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                              <Pill fg={st.fg} bg={st.bg} stamp>{st.label}</Pill>
                              {aguardandoInicioRow && <span title="Lembrete: iniciar serviço" style={{ width: 7, height: 7, borderRadius: "50%", background: COLORS.red, display: "inline-block" }} />}
                            </div>
                          </td>
                          <td style={{ padding: "11px 16px", fontSize: 12, color: COLORS.steel, borderBottom: `1px solid ${COLORS.border}` }}>{fmtDate(p.dataProtocolo)}</td>
                          <td style={{ padding: "11px 16px", fontSize: 12, color: COLORS.steel, borderBottom: `1px solid ${COLORS.border}` }}>{fmtDate(p.dataPrevisaoOrgao)}</td>
                          <td style={{ padding: "11px 16px", borderBottom: `1px solid ${COLORS.border}` }}><Pill fg={prazo.fg} bg={prazo.bg}>{prazo.label}</Pill></td>
                          <td style={{ padding: "11px 16px", fontSize: 12, borderBottom: `1px solid ${COLORS.border}`, color: ds !== null && ds > 15 ? COLORS.red : COLORS.steel }}>{ds !== null ? `há ${ds}d` : "—"}</td>
                          <td style={{ padding: "11px 16px", borderBottom: `1px solid ${COLORS.border}` }}><ChevronRight size={15} color={COLORS.steel} /></td>
                        </tr>
                      );
                    })}
                    {buscados.length === 0 && <tr><td colSpan={11} style={{ padding: 30, textAlign: "center", color: COLORS.steel, fontSize: 13 }}>Nenhum processo encontrado com esses filtros.</td></tr>}
                  </tbody>
                </table>
              </div>
              <Pagination page={processosPage} setPage={setProcessosPage} pageSize={processosPageSize} setPageSize={setProcessosPageSize} totalItems={buscados.length} />
            </div>
          </>
        )}

        {tab === "clientes" && <ClientesPage contratos={contratos} onAddContrato={(row) => addContratoManual(row, false)} isAdmin={isAdmin} />}
        {tab === "unidades" && <UnidadesPage contratos={contratos} onAddContrato={(row) => addContratoManual(row, false)} />}
        {tab === "contratos" && <ContratosPage contratos={contratos} onUpdateContrato={updateContrato} onAddContrato={(row) => addContratoManual(row, true)} onDeleteContrato={deleteContrato} isAdmin={isAdmin} />}
        {tab === "dashboard-planejamento" && isAdmin && <PlanejamentoMensalPage contratos={contratos} />}
        {tab === "importar-contratos" && <ImportarClientesContratosPage onImport={importarContratosPersistindo} />}

        {tab === "atualizacoes" && <AtualizacoesPage processos={processos} onOpenProcesso={(p) => setSelected(p)} />}
        {tab === "acessos" && isAdmin && <GerenciarAcessosPage usuarioLogado={usuarioLogado} />}
      </main>
      </div>

      <footer style={{ borderTop: `1px solid ${COLORS.border}`, background: COLORS.panel, padding: "10px 28px", textAlign: "right", flexShrink: 0 }}>
        <div style={{ fontSize: 11.5, color: COLORS.steelLight, fontWeight: 600 }}>Supervisor - Felipe Moura</div>
        <div style={{ fontSize: 10.5, color: COLORS.steel, marginTop: 2, letterSpacing: "0.03em" }}>Controle Operacional e Financeiro</div>
      </footer>

      {selected && <DetailModal processo={selected} processos={processos} onClose={() => setSelected(null)} onUpdate={(novo) => { updateProcesso(novo); setSelected(novo); }} onOpenProcesso={(p) => setSelected(p)} onConcluir={concluirProcesso} />}
      {showNew && <NewProcessModal processos={processos} onClose={() => setShowNew(false)} onSave={async (p) => {
        const { id, ...campos } = processoToRow(p);
        const { data } = await supabase.from("processos").insert(campos).select().single();
        setProcessos((prev) => [data ? rowToProcesso(data) : p, ...prev]);
      }} isAdmin={isAdmin} />}
    </div>
  );
}

/* ============================================================
   APP — controla a autenticação (protótipo de testes internos)
   ============================================================ */
export default function App() {
  const [usuarioLogado, setUsuarioLogado] = useState(null);
  const [verificandoSessao, setVerificandoSessao] = useState(true);

  useEffect(() => {
    if (!SUPABASE_CONFIGURADO) { setVerificandoSessao(false); return; }
    let ativo = true;
    async function carregarPerfilDaSessao(userId) {
      const { data: perfil } = await supabase.from("profiles").select("id, usuario, nome, role").eq("id", userId).single();
      if (ativo) setUsuarioLogado(perfil || null);
    }
    supabase.auth.getSession().then(({ data }) => {
      if (data.session?.user) carregarPerfilDaSessao(data.session.user.id);
      if (ativo) setVerificandoSessao(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      if (!session?.user) { setUsuarioLogado(null); return; }
      carregarPerfilDaSessao(session.user.id);
    });
    return () => { ativo = false; listener.subscription.unsubscribe(); };
  }, []);

  const sair = async () => { await supabase.auth.signOut(); setUsuarioLogado(null); };

  if (verificandoSessao) {
    return <div style={{ minHeight: "100vh", background: COLORS.bg, display: "flex", alignItems: "center", justifyContent: "center", color: COLORS.steel, fontFamily: "'Inter', sans-serif", fontSize: 13 }}>Carregando...</div>;
  }
  if (!usuarioLogado) return <LoginScreen onLogin={setUsuarioLogado} />;
  return <ControleProcessos usuarioLogado={usuarioLogado} onLogout={sair} />;
}
