/* =====================================================================
   COLORÍMETRO DS-01 — dentro do Gerador-OS (19/10/2026)
   =====================================================================
   (Junior: "Copie o programa Colorímetro na pasta C:\Users\Pichau\Desktop\
   Colorímetro e integre-o ao programa Gerador OS, em item na barra lateral
   com nome Colorímetro".)

   O original é um programa Python (colorimetro/DS01-PC/, copiado inteiro
   para cá): fala com o CHNSpec DS-01 por Bluetooth LE pela biblioteca bleak.
   No navegador o caminho é o Web Bluetooth do Chrome — o mesmo que a
   ponte.html do original já usava no celular. Por isso esta tela é o
   protocolo.py e o colorimetro.py reescritos, com as MESMAS regras:
     · serviço FFE0 / característica FFE1, o byte 0x01 que acorda o aparelho
       antes de cada comando, pacote BB <cmd> ... FF <soma % 256>;
     · espectro → XYZ → L*a*b* com as tabelas extraídas do app oficial
       (colorimetro/iluminantes.json) e as mesmas constantes;
     · dE*ab, dE00, dEcmc e dE94; aprova só pelo dE escolhido; as dicas por
       coluna saem dos valores já arredondados na tela, como no app.

   ONDE FICAM AS MEDIÇÕES: neste navegador (IndexedDB), uma lista por dia —
   o equivalente da pasta "medicoes" do original, que também era por
   computador. Não vão para o banco compartilhado de propósito: cada
   medição leva a curva espectral inteira, e o blob único do programa é
   baixado inteiro por todo mundo (ver o egress do Supabase). Para levar a
   outro lugar: Exportar CSV (o mesmo formato do original, abre no Excel);
   o Importar CSV traz os arquivos da pasta "medicoes" do programa antigo.

   Arquivo separado do app.js pelo mesmo motivo do assistente.js: o app.js
   tem 2,4 MB e a tela não precisa de nada dele além da navegação. A página
   é a section[data-page="colorimetro"] do index.html; ela é desenhada quando
   aparece (MutationObserver na classe hidden), sem gancho no goto.
   ===================================================================== */
(function () {
  'use strict';

  const SERVICO = 0xffe0, CARAC = 0xffe1;
  const NOMES_CONHECIDOS = ['CM', 'NJ', 'GJ', 'CH', 'DS'];
  const MODOS = { 0: 'SCI', 1: 'SCE', 3: 'M0', 4: 'M1', 5: 'M2', 6: 'M3' };
  const FORMULAS = ['dE*ab', 'dE00', 'dEcmc', 'dE94'];
  const CASAS = 1;   // o app mostra 1 casa para o DS-01
  const CHAVE_CFG = 'gos:colorimetro:cfg';

  // Os valores do app no celular (capturas de 27/08/2026), que são também os do
  // config.json do programa no PC em 19/10/2026 — conferidos um a um, tolerâncias
  // inclusive (Junior: "copiar as configurações de tolerância como o programa
  // está configurado").
  const CFG_PADRAO = {
    iluminante1: 'D65/10', iluminante2: 'A/2', modo: '1', formula: 'dE*ab', media: 3,
    cmc_l: 2, cmc_c: 1,
    tol: { dL: 0.4, da: 0.1, db: 0.5, 'dE*ab': 0.6, dC: 0.6, dH: 0.6, "dL'": 0.6, "dC'": 0.6, "dH'": 0.6,
      dE00: 0.6, dEcmc: 0.6, dE94: 0.6 },
    nome_padrao: 'Padrão', nome_amostra: 'Amostra', k94: [1, 1, 1], k00: [1, 1, 1],
    botao: 'amostra', ultimo_aparelho: null, versao_cfg: 2
  };
  /* PADRÃO DA FÁBRICA EM TODOS OS COMPUTADORES (30/10/2026, Jones: "essa configuração
     seja parte do programa para que qualquer computador utilize"). Média de 3
     leituras (uma leitura fora, como a Amostra 7 das 11:59, não decide sozinha) e
     modo SCE (no "Automático" o DS-01 mede em SCI). Computador que já tinha
     configurações salvas de antes (versao_cfg < 2) recebe esses dois valores uma
     vez; depois disso, o que alguém mudar em Configurações vale naquele computador. */
  const VERSAO_CFG = 2;

  const esperar = ms => new Promise(r => setTimeout(r, ms));
  const escH = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const dois = n => String(n).padStart(2, '0');
  const diaIso = d => `${d.getFullYear()}-${dois(d.getMonth() + 1)}-${dois(d.getDate())}`;
  const horaTxt = d => `${dois(d.getHours())}:${dois(d.getMinutes())}:${dois(d.getSeconds())}`;
  const dataTxt = d => `${dois(d.getDate())}/${dois(d.getMonth() + 1)}/${d.getFullYear()}`;
  const hex = u8 => Array.from(u8, b => b.toString(16).padStart(2, '0')).join(' ');

  function carregarCfg() {
    const cfg = JSON.parse(JSON.stringify(CFG_PADRAO));
    try {
      const salvo = JSON.parse(localStorage.getItem(CHAVE_CFG) || 'null');
      if (salvo && typeof salvo === 'object') {
        const tol = Object.assign({}, cfg.tol, salvo.tol || {});
        Object.assign(cfg, salvo);
        cfg.tol = tol;
        if (!(salvo.versao_cfg >= VERSAO_CFG)) {
          cfg.media = CFG_PADRAO.media;
          cfg.modo = CFG_PADRAO.modo;
          cfg.versao_cfg = VERSAO_CFG;
          try { localStorage.setItem(CHAVE_CFG, JSON.stringify(cfg)); } catch (e) { /* sem storage */ }
        }
      }
    } catch (e) { /* sem storage: vale o padrão */ }
    return cfg;
  }
  function salvarCfg() {
    try { localStorage.setItem(CHAVE_CFG, JSON.stringify(cfg)); } catch (e) { /* sem storage */ }
  }
  let cfg = carregarCfg();

  /* ---------------- números ---------------- */
  // Arredonda como o app (toFixed do JavaScript: o valor exato, metade se afasta do zero).
  const arred = (v, c = CASAS) => Number(Number(v).toFixed(c));
  const num = (v, c = CASAS) => arred(v, c).toFixed(c).replace('.', ',');
  const numTol = v => String(v).replace('.', ',');
  const dentro = (v, lim) => Math.abs(v) <= lim + 1e-9;   // 0,4000000001 não passa do limite de 0,4
  const lerNum = s => { const v = parseFloat(String(s).replace(',', '.')); return isNaN(v) ? null : v; };

  /* ---------------- cor ---------------- */
  let T = null;   // iluminantes.json
  async function carregarTabelas() {
    if (T) return T;
    const r = await fetch('colorimetro/iluminantes.json', { cache: 'force-cache' });
    if (!r.ok) throw new Error('não consegui baixar as tabelas de cor (colorimetro/iluminantes.json)');
    T = await r.json();
    return T;
  }

  const rad = d => d * Math.PI / 180;
  const grau = r => r * 180 / Math.PI;
  const mod360 = x => ((x % 360) + 360) % 360;
  const lch = (L, a, b) => [Math.hypot(a, b), mod360(grau(Math.atan2(b, a)))];

  // Alinha o espectro na grade 360–780 nm (43 pontos), como fillSpectralArray do app.
  function completar(esp, ws) {
    if (esp.length === T.count) return esp.slice();
    const ini = Math.floor(Math.abs(T.wave_start - ws) / 10);
    const d = new Array(ini).fill(0).concat(esp);
    while (d.length < T.count) d.push(0);
    return d.slice(0, T.count);
  }
  function xyz(e, il) {
    const w = T.illuminants[il].weights;
    return [0, 1, 2].map(c => w[c].reduce((s, wi, i) => s + wi * e[i], 0));
  }
  // Mesma fórmula (e constantes) de XYZtoCIE_LabCH do app.
  function lab(x, il) {
    const [xn, yn, zn] = T.illuminants[il].white;
    const f = t => t > 0.008856 ? Math.pow(t, 0.3333333) : 7.787 * t + 0.138;
    const ty = x[1] / yn;
    const L = ty > 0.008856 ? 116 * Math.pow(ty, 0.3333333) - 16 : 903.3 * ty;
    return [Math.max(L, 0), 500 * (f(x[0] / xn) - f(ty)), 200 * (f(ty) - f(x[2] / zn))];
  }
  const espectroParaLab = (esp, ws, il) => lab(xyz(completar(esp, ws), il), il);

  const deAb = (l1, l2) => Math.hypot(l1[0] - l2[0], l1[1] - l2[1], l1[2] - l2[2]);

  function de2000(l1, l2, kL = 1, kC = 1, kH = 1) {
    const [L1, a1, b1] = l1, [L2, a2, b2] = l2;
    const C1 = Math.hypot(a1, b1), C2 = Math.hypot(a2, b2);
    const Cm = (C1 + C2) / 2;
    const G = 0.5 * (1 - Math.sqrt(Cm ** 7 / (Cm ** 7 + 25 ** 7)));
    const a1p = (1 + G) * a1, a2p = (1 + G) * a2;
    const C1p = Math.hypot(a1p, b1), C2p = Math.hypot(a2p, b2);
    const h1p = C1p ? mod360(grau(Math.atan2(b1, a1p))) : 0;
    const h2p = C2p ? mod360(grau(Math.atan2(b2, a2p))) : 0;
    const dLp = L2 - L1, dCp = C2p - C1p;
    let dhp;
    if (C1p * C2p === 0) dhp = 0;
    else if (Math.abs(h2p - h1p) <= 180) dhp = h2p - h1p;
    else if (h2p - h1p > 180) dhp = h2p - h1p - 360;
    else dhp = h2p - h1p + 360;
    const dHp = 2 * Math.sqrt(C1p * C2p) * Math.sin(rad(dhp / 2));
    const Lm = (L1 + L2) / 2, Cmp = (C1p + C2p) / 2;
    let hm;
    if (C1p * C2p === 0) hm = h1p + h2p;
    else if (Math.abs(h1p - h2p) <= 180) hm = (h1p + h2p) / 2;
    else if (h1p + h2p < 360) hm = (h1p + h2p + 360) / 2;
    else hm = (h1p + h2p - 360) / 2;
    const Tt = 1 - 0.17 * Math.cos(rad(hm - 30)) + 0.24 * Math.cos(rad(2 * hm))
      + 0.32 * Math.cos(rad(3 * hm + 6)) - 0.20 * Math.cos(rad(4 * hm - 63));
    const dtheta = 30 * Math.exp(-(((hm - 275) / 25) ** 2));
    const Rc = 2 * Math.sqrt(Cmp ** 7 / (Cmp ** 7 + 25 ** 7));
    const Sl = 1 + 0.015 * (Lm - 50) ** 2 / Math.sqrt(20 + (Lm - 50) ** 2);
    const Sc = 1 + 0.045 * Cmp, Sh = 1 + 0.015 * Cmp * Tt;
    const Rt = -Math.sin(rad(2 * dtheta)) * Rc;
    return Math.sqrt((dLp / (kL * Sl)) ** 2 + (dCp / (kC * Sc)) ** 2 + (dHp / (kH * Sh)) ** 2
      + Rt * (dCp / (kC * Sc)) * (dHp / (kH * Sh)));
  }

  function de94(ref, amo, kL = 1, kC = 1, kH = 1) {
    const [L1, a1, b1] = ref, [L2, a2, b2] = amo;
    const C1 = Math.hypot(a1, b1), C2 = Math.hypot(a2, b2);
    const dL = L1 - L2, dC = C1 - C2;
    const dH2 = Math.max(0, (a1 - a2) ** 2 + (b1 - b2) ** 2 - dC ** 2);
    const Sc = 1 + 0.045 * C1, Sh = 1 + 0.015 * C1;
    return Math.sqrt((dL / kL) ** 2 + (dC / (kC * Sc)) ** 2 + dH2 / (kH * Sh) ** 2);
  }

  function deCmc(ref, amo, l = 2, c = 1) {
    const [L1, a1, b1] = ref, [L2, a2, b2] = amo;
    const C1 = Math.hypot(a1, b1), C2 = Math.hypot(a2, b2);
    const dL = L1 - L2, dC = C1 - C2;
    const dH2 = Math.max(0, (a1 - a2) ** 2 + (b1 - b2) ** 2 - dC ** 2);
    const h1 = mod360(grau(Math.atan2(b1, a1)));
    const F = Math.sqrt(C1 ** 4 / (C1 ** 4 + 1900));
    const Tt = (h1 >= 164 && h1 <= 345) ? 0.56 + Math.abs(0.2 * Math.cos(rad(h1 + 168)))
      : 0.36 + Math.abs(0.4 * Math.cos(rad(h1 + 35)));
    const Sl = L1 < 16 ? 0.511 : 0.040975 * L1 / (1 + 0.01765 * L1);
    const Sc = 0.0638 * C1 / (1 + 0.0131 * C1) + 0.638;
    const Sh = Sc * (F * Tt + 1 - F);
    return Math.sqrt((dL / (l * Sl)) ** 2 + (dC / (c * Sc)) ** 2 + dH2 / Sh ** 2);
  }

  // Cor aproximada para a tela (Lab D65 → sRGB).
  function labHex(L, a, b) {
    const fy = (L + 16) / 116, fx = fy + a / 500, fz = fy - b / 200;
    const inv = t => t ** 3 > 0.008856 ? t ** 3 : (t - 16 / 116) / 7.787;
    const X = 0.95047 * inv(fx), Y = inv(fy), Z = 1.08883 * inv(fz);
    const r = 3.2406 * X - 1.5372 * Y - 0.4986 * Z;
    const g = -0.9689 * X + 1.8758 * Y + 0.0415 * Z;
    const bl = 0.0557 * X - 0.2040 * Y + 1.0570 * Z;
    const gama = c => { c = Math.max(0, Math.min(1, c)); return c <= 0.0031308 ? 12.92 * c : 1.055 * c ** (1 / 2.4) - 0.055; };
    return '#' + [r, g, bl].map(c => Math.round(gama(c) * 255).toString(16).padStart(2, '0')).join('');
  }

  /* ---------------- protocolo do aparelho ---------------- */
  class ErroAparelho extends Error {}

  function pacote(...bytes) {
    const b = bytes.map(x => x & 0xFF);
    b.push(b.reduce((s, x) => s + x, 0) % 256);
    return Uint8Array.from(b);
  }
  const checksumOk = d => d.length >= 2 && d.slice(0, -1).reduce((s, x) => s + x, 0) % 256 === d[d.length - 1];
  function texto(buf) {
    const o = [];
    for (const c of buf) { if (c === 0) break; o.push(c); }
    return new TextDecoder().decode(Uint8Array.from(o)).trim();
  }
  const f32 = (d, off) => new DataView(d.buffer, d.byteOffset, d.byteLength).getFloat32(off, true);
  const u32 = (d, off) => new DataView(d.buffer, d.byteOffset, d.byteLength).getUint32(off, true);
  const juntar = (a, b) => { const n = new Uint8Array(a.length + b.length); n.set(a); n.set(b, a.length); return n; };

  // Regra do app: ColorMeter SE (código 776) com firmware >= 20210420.
  function protocoloNovo(info) {
    const d = parseInt(String(info.firmware || '').split('.').pop(), 10);
    return info.codigo === 776 && !isNaN(d) && d >= 20210420;
  }
  const modoPadrao = info => [778, 779].includes(info.codigo) ? 3 : 0;

  class DS01 {
    constructor(t, aoBotao, log) {
      this.t = t;
      t.aoReceber = d => this._receber(d);
      this.aoBotao = aoBotao;
      this.log = log;
      this.info = null;
      this.buf = new Uint8Array(0);
      this.esperando = false;
      this.acordar = null;
      this.fila = Promise.resolve();
    }
    _receber(d) {
      d = Uint8Array.from(d);
      if (this.esperando) {
        this.buf = juntar(this.buf, d);
        if (this.acordar) this.acordar();
        return;
      }
      // Notificação espontânea: o botão de medir do aparelho foi apertado.
      if (d.length >= 10 && d[0] === 0xBB && d[1] === 0x01 && checksumOk(d.slice(0, 10))) {
        if (this.aoBotao) this.aoBotao(d[2], d[3]);
      } else {
        this.log('recebido fora de hora: ' + hex(d));
      }
    }
    async _escrever(d) {
      for (let i = 0; i < d.length; i += 20) {
        await this.t.escrever(d.slice(i, i + 20));
        await esperar(10);
      }
    }
    // A resposta começa em BB <cmd>; o que vier antes é sobra de comando anterior.
    _resposta(cmd) {
      for (let i = 0; i < this.buf.length - 1; i++) if (this.buf[i] === 0xBB && this.buf[i + 1] === cmd) return this.buf.slice(i);
      return new Uint8Array(0);
    }
    _emFila(fn) {
      const p = this.fila.then(fn, fn);
      this.fila = p.catch(() => {});
      return p;
    }
    async comando(quadro, tamanho = 10, timeout = 3000) {
      // Leituras podem ser repetidas sem efeito colateral; medir e calibrar não.
      const tentativas = [0x01, 0x10, 0x11].includes(quadro[1]) ? 1 : 2;
      for (let k = 1; k <= tentativas; k++) {
        try {
          return await this._emFila(() => this._comando(quadro, tamanho, timeout));
        } catch (e) {
          this.log(`${e.message} (tentativa ${k}); bytes recebidos: ${hex(this.buf) || 'nenhum'}`);
          if (!(e instanceof ErroAparelho) || k === tentativas) throw e;
          await esperar(300);
        }
      }
    }
    async _comando(quadro, tamanho, timeout) {
      const cmd = quadro[1];
      let data;
      this.buf = new Uint8Array(0);
      this.esperando = true;
      try {
        await this.t.escrever(Uint8Array.of(0x01));   // acordar
        await esperar(100);
        this.buf = new Uint8Array(0);
        await this._escrever(quadro);
        const fim = Date.now() + timeout;
        while (this._resposta(cmd).length < tamanho) {
          const resto = fim - Date.now();
          if (resto <= 0) break;
          await new Promise(r => { this.acordar = r; setTimeout(r, resto); });
          this.acordar = null;
        }
        await esperar(50);   // sobras do último pacote
        data = this._resposta(cmd);
      } finally {
        this.esperando = false;
      }
      if (data.length < tamanho) throw new ErroAparelho(`resposta incompleta (${data.length}/${tamanho} bytes) ao comando 0x${cmd.toString(16)}`);
      const ns = [...new Set([data.length, tamanho, 10, 20, 200])].filter(n => n <= data.length);
      if (!ns.some(n => checksumOk(data.slice(0, n)))) throw new ErroAparelho(`checksum inválido na resposta ao comando 0x${cmd.toString(16)}`);
      return data;
    }
    async lerInfo() {
      const d = await this.comando(pacote(0xBB, 0x12, 0x01, 0, 0, 0, 0, 0, 0xFF), 200);
      this.info = { codigo: d[5] + (d[6] << 8), modelo: texto(d.slice(37, 67)), serie: texto(d.slice(67, 97)),
        firmware: texto(d.slice(97, 127)), hardware: texto(d.slice(127, 157)) };
      return this.info;
    }
    async lerCalibracao() {
      let d = await this.comando(pacote(0xBB, 0x1E, 0, 0, 0, 0, 0, 0, 0xFF), 10);
      if (d.length < 12) d = juntar(d, new Uint8Array(12 - d.length));
      return { branco_ok: d[2] === 1, branco_data: u32(d, 3), preto_ok: d[7] === 1, preto_data: u32(d, 8) };
    }
    async calibrar(tipo) {
      const cmd = tipo === 'preto' ? 0x10 : 0x11;
      const d = await this.comando(pacote(0xBB, cmd, 1, 0, 0, 0, 0, 0, 0xFF), 10, 12000);
      return d[2] === 0;
    }
    async medir(modo) {
      if (modo == null) modo = this.info ? modoPadrao(this.info) : 0;
      const al = Array.from({ length: 4 }, () => Math.floor(Math.random() * 256));
      const d = await this.comando(pacote(0xBB, 0x01, modo, ...al, 0, 0xFF), 10, 6000);
      if (d[3] !== 0) throw new ErroAparelho(`o aparelho retornou erro ${d[3]} na medição`);
      return this.baixarMedicao(d[2]);
    }
    async baixarMedicao(modo) {
      const cmd = (this.info && protocoloNovo(this.info)) ? 0x05 : 0x02;
      const d = await this.comando(pacote(0xBB, cmd, modo + 0x10, 0, 0, 0, 0, 0, 0xFF), 200);
      const m = { modo: d[2] - 0x10, wave_start: (d[4] << 8) + d[5], intervalo: d[6], espectro: [], lab_aparelho: null };
      const tipo = d[3], n = d[7];
      if (tipo === 0) m.espectro = Array.from({ length: n }, (_, i) => Math.max(0, f32(d, 8 + 4 * i)));
      else m.lab_aparelho = [184, 188, 192].map(o => f32(d, o));
      return m;
    }
  }

  // Conexão real, pelo Web Bluetooth do Chrome.
  class TransporteBLE {
    constructor() { this.aoReceber = null; this.aoDesconectar = null; this.c = null; this.dev = null; }
    async conectar(dev) {
      this.dev = dev;
      this._caiu = () => { this.c = null; if (this.aoDesconectar) this.aoDesconectar(); };
      dev.addEventListener('gattserverdisconnected', this._caiu);
      // No Windows a 1ª conexão às vezes cai logo depois de conectar: tenta de novo.
      for (let k = 1; k <= 4; k++) {
        try {
          const srv = await dev.gatt.connect();
          const s = await srv.getPrimaryService(SERVICO);
          this.c = await s.getCharacteristic(CARAC);
          this._ouvir = e => {
            const v = e.target.value;
            if (this.aoReceber) this.aoReceber(new Uint8Array(v.buffer.slice(v.byteOffset, v.byteOffset + v.byteLength)));
          };
          this.c.addEventListener('characteristicvaluechanged', this._ouvir);
          await this.c.startNotifications();
          return;
        } catch (e) {
          try { if (dev.gatt.connected) dev.gatt.disconnect(); } catch (_) { /* já caiu */ }
          if (k === 4) throw e;
          await esperar(1500);
        }
      }
    }
    async escrever(d) {
      if (!this.c) throw new ErroAparelho('colorímetro desconectado');
      if (this.c.writeValueWithoutResponse) await this.c.writeValueWithoutResponse(d);
      else await this.c.writeValue(d);
    }
    async reassinar() {
      try { await this.c.stopNotifications(); } catch (e) { /* não estava */ }
      await this.c.startNotifications();
    }
    async desconectar() {
      if (this.dev) {
        this.dev.removeEventListener('gattserverdisconnected', this._caiu);
        try { if (this.dev.gatt.connected) this.dev.gatt.disconnect(); } catch (e) { /* já caiu */ }
      }
      this.c = null;
    }
  }

  // Aparelho falso que responde ao mesmo protocolo — para conhecer a tela sem o DS-01.
  class TransporteSimulado {
    constructor() {
      this.aoReceber = null; this.aoDesconectar = null;
      const agora = Math.floor(Date.now() / 1000) - 3600;
      this.cal = { 0x10: agora, 0x11: agora };
    }
    async conectar() { await esperar(300); }
    async desconectar() {}
    async reassinar() {}
    _responder(data, atraso = 50) {
      setTimeout(async () => {
        for (let i = 0; i < data.length; i += 20) {   // fragmentação do BLE
          if (this.aoReceber) this.aoReceber(data.slice(i, i + 20));
          await esperar(5);
        }
      }, atraso);
    }
    _fechar(arr) { arr[arr.length - 1] = arr.slice(0, -1).reduce((s, x) => s + x, 0) % 256; return arr; }
    apertarBotao(modo = 0) { this._responder(pacote(0xBB, 0x01, modo, 0, 0, 0, 0, 0, 0xFF)); }
    async escrever(d) {
      if (d.length < 9 || d[0] !== 0xBB) return;
      const cmd = d[1];
      if (cmd === 0x12) {
        const r = new Uint8Array(200);
        r[0] = 0xBB; r[1] = 0x12; r[5] = 774 & 0xFF; r[6] = 774 >> 8;
        [[37, 'DS-01'], [67, 'SIM0000001'], [97, 'V1.0.20240101'], [127, 'H1.0']]
          .forEach(([o, s]) => r.set(new TextEncoder().encode(s), o));
        r[198] = 0xFF;
        this._responder(this._fechar(r));
      } else if (cmd === 0x1E) {
        const r = new Uint8Array(20);
        r[0] = 0xBB; r[1] = 0x1E; r[2] = 1;
        new DataView(r.buffer).setUint32(3, this.cal[0x11], true);
        r[7] = 1;
        new DataView(r.buffer).setUint32(8, this.cal[0x10], true);
        r[18] = 0xFF;
        this._responder(this._fechar(r));
      } else if (cmd === 0x10 || cmd === 0x11) {
        this.cal[cmd] = Math.floor(Date.now() / 1000);
        this._responder(pacote(0xBB, cmd, 0, 0, 0, 0, 0, 0, 0xFF), 1500);
      } else if (cmd === 0x01) {
        this._responder(pacote(0xBB, 0x01, d[2], 0, 0x01, 0x90, 10, 31, 0xFF), 800);
      } else if (cmd === 0x02 || cmd === 0x05) {
        const r = new Uint8Array(200);
        r[0] = 0xBB; r[1] = cmd; r[2] = d[2]; r[3] = 0;
        r[4] = 0x01; r[5] = 0x90; r[6] = 10; r[7] = 31;   // 400 nm, 10 nm, 31 pontos
        const dv = new DataView(r.buffer);
        const c = 450 + Math.random() * 200, base = 5 + Math.random() * 25, amp = 20 + Math.random() * 40;
        for (let i = 0; i < 31; i++) {
          const wl = 400 + 10 * i;
          dv.setFloat32(8 + 4 * i, base + amp * Math.exp(-(((wl - c) / (50 + Math.random() * 40)) ** 2)), true);
        }
        r[198] = 0xFF;
        this._responder(this._fechar(r));
      }
    }
  }

  /* ---------------- medições guardadas (IndexedDB, por navegador) ---------------- */
  const BD = {
    db: null,
    abrir() {
      if (this.db) return Promise.resolve(this.db);
      return new Promise((ok, erro) => {
        let r;
        try { r = indexedDB.open('gos-colorimetro', 1); } catch (e) { erro(e); return; }
        r.onupgradeneeded = () => {
          const s = r.result.createObjectStore('medicoes', { keyPath: 'id' });
          s.createIndex('dia', 'dia');
        };
        r.onsuccess = () => { this.db = r.result; ok(this.db); };
        r.onerror = () => erro(r.error);
      });
    },
    async _tx(modo, fn) {
      const db = await this.abrir();
      return new Promise((ok, erro) => {
        const tx = db.transaction('medicoes', modo);
        const req = fn(tx.objectStore('medicoes'));
        tx.oncomplete = () => ok(req && req.result);
        tx.onerror = () => erro(tx.error);
      });
    },
    gravar(reg) { return this._tx('readwrite', s => s.put(reg)); },
    todos() { return this._tx('readonly', s => s.getAll()); },
    doDia(dia) { return this._tx('readonly', s => s.index('dia').getAll(dia)); }
  };

  /* ---------------- estado ---------------- */
  const S = {
    t: null, dev: null, cal: null, ocupado: false, conectando: false,
    padrao: null, leituras: [], acumulando: [], cont: { padrao: 0, amostra: 0 },
    atual: null, pagina: 0, status: 'Desconectado', aviso: '', log: [],
    pedido: null, janela: null, os: null, av: null, relSalvo: null, histDias: [], histDia: '', histLidas: [], semBD: false
  };
  let seq = 0;
  const novoId = () => 'cm_' + Date.now().toString(36) + '_' + (++seq);

  function log(t) {
    S.log.unshift(new Date().toLocaleTimeString('pt-BR') + '  ' + t);
    S.log = S.log.slice(0, 80);
  }
  function status(t) { S.status = t; desenhar(); }

  function novaLeitura(dados) {
    const l = Object.assign({ id: novoId(), refId: null }, dados);
    recalcular(l);
    return l;
  }
  function recalcular(l) {
    if (l.espectro && l.espectro.length) {
      l.lab1 = espectroParaLab(l.espectro, l.wave_start, cfg.iluminante1);
      l.lab2 = espectroParaLab(l.espectro, l.wave_start, cfg.iluminante2);
    } else {
      l.lab1 = l.lab_aparelho; l.lab2 = null;
    }
  }

  function comparar(ref, amo) {
    const tol = cfg.tol;
    const l1 = ref.lab1, l2 = amo.lab1;
    const [dL, da, db] = [0, 1, 2].map(i => l2[i] - l1[i]);
    const dC = lch(...l2)[0] - lch(...l1)[0];
    const dEab = deAb(l1, l2);
    const dH = Math.sign(l1[1] * l2[2] - l2[1] * l1[2] || 1) * Math.sqrt(Math.max(0, dEab ** 2 - dL ** 2 - dC ** 2));
    const des = { 'dE*ab': dEab, dE00: de2000(l1, l2, ...cfg.k00), dEcmc: deCmc(l1, l2, cfg.cmc_l, cfg.cmc_c), dE94: de94(l1, l2, ...cfg.k94) };
    /* APROVAÇÃO TAMBÉM PELAS TOLERÂNCIAS dL*, da*, db* (30/10/2026, Jones: "Sim" à
       pergunta se a aprovação devia considerar dL*, da* e db*). Antes aprovava só
       pelo dE escolhido, como o app do aparelho: a Amostra 1 com dL* +0,5 (limite
       ±0,4) saía "Aprovado". Agora só aprova se o dE estiver dentro E as três
       colunas estiverem "Aprovado". O dE usa o valor exato.
       SEM ARREDONDAR (12/11/2026, Junior: "não deve haver aproximação"; caso OS
       0715): as colunas usavam os valores arredondados a 1 casa, e o
       arredondamento sozinho deslocava até 0,1 — o tamanho inteiro da
       tolerância do da*. Na 0715 isso pôs no mesmo tom amostras com dL* real de
       0,47 e da* de 0,16 (limites 0,4 e 0,1), que a olho eram diferentes. Agora
       a decisão usa a diferença REAL, e a diferença aparece com 2 casas.
       Referência (30/10/2026, 11:48): par de malha preta que NÃO é compatível a
       olho nu deu dE*ab 0,2–0,4, mas a* sempre mais verde (-0,1 a -0,3): quem
       reprova esse par é o da* ±0,1. Por isso as tolerâncias ficam como estão. */
    const deOk = des[cfg.formula] <= tol[cfg.formula];
    const d1 = [0, 1, 2].map(i => l2[i] - l1[i]);
    const rot = [['Claro', 'Escuro'], ['Vermelho', 'Verde'], ['Amarelo', 'Azul']];
    const txt = [['mais clara', 'mais escura'], ['mais vermelha', 'mais verde'], ['mais amarela', 'mais azul']];
    const dicas = [], desc = [];
    ['dL', 'da', 'db'].forEach((k, i) => {
      const v = d1[i];
      if (dentro(v, tol[k])) dicas.push(['Aprovado', true]);
      else { dicas.push([v > 0 ? rot[i][0] : rot[i][1], false]); desc.push(v > 0 ? txt[i][0] : txt[i][1]); }
    });
    const aprovado = deOk && dicas.every(x => x[1]);
    return { d: [dL, da, db, dC, dH], d1, dicas, des, aprovado, deOk, desc };
  }

  const modoAtual = () => cfg.modo !== 'auto' ? parseInt(cfg.modo, 10) : (S.dev && S.dev.info ? modoPadrao(S.dev.info) : 0);
  const refDe = l => l.refId ? S.leituras.find(x => x.id === l.refId) || (S.padrao && S.padrao.id === l.refId ? S.padrao : null) : null;

  /* ---------------- conexão ---------------- */
  function semBluetooth() {
    if (navigator.bluetooth) return '';
    if (!window.isSecureContext) return 'O Bluetooth do navegador só funciona em endereço https:// (abra o programa por https://192.168.1.50).';
    return 'Este navegador não tem Web Bluetooth. Use o Google Chrome (ou o Edge) no computador.';
  }

  async function conectar(simular) {
    if (S.ocupado || S.conectando) return;
    try { await carregarTabelas(); } catch (e) { S.aviso = e.message; desenhar(); return; }
    let t, alvoNome;
    if (simular) {
      t = new TransporteSimulado();
      alvoNome = 'CM-DS01 (simulado)';
    } else {
      const motivo = semBluetooth();
      if (motivo) { S.aviso = motivo; desenhar(); return; }
      let dev = null;
      try {
        // A escolha do aparelho é do Chrome (a janelinha com a lista): precisa do clique.
        dev = await navigator.bluetooth.requestDevice({
          filters: NOMES_CONHECIDOS.map(n => ({ namePrefix: n })).concat([{ services: [SERVICO] }]),
          optionalServices: [SERVICO]
        });
      } catch (e) {
        S.aviso = /cancel/i.test(e.message) ? 'Nenhum aparelho escolhido.'
          : /adapter|not available|bluetooth/i.test(e.message) ? 'Não encontrei um adaptador Bluetooth ligado neste computador. Confira em Configurações › Dispositivos › Bluetooth.'
          : 'Não deu para procurar o aparelho: ' + e.message;
        desenhar(); return;
      }
      t = new TransporteBLE();
      t._dev = dev;
      alvoNome = dev.name || 'colorímetro';
    }
    S.conectando = true; S.aviso = '';
    status(`Conectando a ${alvoNome}…`);
    const dev = new DS01(t, aoBotao, log);
    t.aoDesconectar = caiu;
    try {
      await t.conectar(t._dev);
      log('conectado a ' + alvoNome);
      await esperar(500);
      let info = null;
      for (let k = 1; k <= 3; k++) {
        try { info = await dev.lerInfo(); break; } catch (e) {
          // Às vezes a assinatura das notificações não "pega": assina de novo.
          log(`leitura de informações falhou (tentativa ${k}): ${e.message}`);
          if (k === 3) throw e;
          await t.reassinar();
          await esperar(500);
        }
      }
      log('aparelho: ' + JSON.stringify(info));
      const cal = await dev.lerCalibracao();
      S.t = t; S.dev = dev; S.cal = cal; S.simulado = !!simular;
      if (!simular) { cfg.ultimo_aparelho = { nome: alvoNome }; salvarCfg(); }
      S.conectando = false;
      S.aviso = (cal.branco_ok && cal.preto_ok) ? '' : 'O aparelho informa que precisa de calibração: use Calibrar branco e Calibrar preto.';
      status(`Conectado a ${alvoNome}`);
    } catch (e) {
      log(`falha ao conectar: ${e.message}`);
      S.conectando = false;
      try { await t.desconectar(); } catch (_) { /* nada */ }
      S.aviso = 'Não foi possível conectar: ' + e.message + '. Confira se o aparelho está ligado, perto e desconectado do celular (o DS-01 aceita uma conexão de cada vez).';
      status('Desconectado');
    }
  }

  async function desconectar() {
    const t = S.t;
    S.t = null; S.dev = null; S.cal = null; S.ocupado = false;
    if (t) { t.aoDesconectar = null; try { await t.desconectar(); } catch (e) { /* nada */ } }
    status('Desconectado');
  }

  function caiu() {
    if (!S.dev) return;
    S.t = null; S.dev = null; S.ocupado = false;
    status('⚠ Conexão perdida — clique em Conectar de novo');
  }

  /* AS PERGUNTAS SÃO UM QUADRO NA PÁGINA (19/10/2026, Junior: "adapte o visual
     ao estilo do Gerador OS... sem janelas pop-up, mas sim janelas normais").
     Calibrar, renomear, limpar e voltar ao padrão abrem este quadro no topo da
     tela, com o texto e os botões; nada de confirm()/prompt() do navegador. */
  /* O QUADRO APARECE ONDE SE CLICOU (20/10/2026, Junior: "corrija a função
     limpar lista na avaliação de amostras, pois o botão não funciona, a lista
     não fica limpa depois de clicar"). A pergunta do Limpar saía no topo da
     tela, e o botão fica embaixo, na coluna da lista — na avaliação de uma OS,
     abaixo ainda do quadro da avaliação. Quem clicava não via a pergunta, e sem
     a resposta a lista não era limpa. Agora o pedido pode dizer o lugar
     (`lugar: 'lista'` = em cima da lista) e a tela rola até ele. */
  function perguntar(pedido) {
    S.pedido = pedido;
    desenhar();
    const q = document.querySelector('#colorimetro-painel .cm-pedido');
    if (q && q.scrollIntoView) q.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
  function htmlPedido() {
    const p = S.pedido;
    return `<div class="card cm-pedido"><div class="card-title">${escH(p.titulo)}</div>
        <div>${escH(p.texto)}</div>
        ${p.campo != null ? `<input id="cm-pedido-valor" class="cm-pedido-campo" value="${escH(p.campo)}" onkeydown="if(event.key==='Enter')Colorimetro.responder(true)">` : ''}
        <div class="cm-botoes"><button type="button" class="btn primary" onclick="Colorimetro.responder(true)">${escH(p.sim)}</button>
        <button type="button" class="btn" onclick="Colorimetro.responder(false)">Cancelar</button></div></div>`;
  }
  function responder(sim) {
    const p = S.pedido;
    S.pedido = null;
    if (!p || !sim) { desenhar(); return; }
    const valor = p.campo != null ? (document.getElementById('cm-pedido-valor') || {}).value : undefined;
    p.executar(valor);
  }

  /* O clique já calibra (20/10/2026, Junior: "não precisa desse painel, pois
     usuário deve clicar no botão e disparar calibração"). O quadro de
     confirmação saiu; a instrução da placa ficou na dica do botão. */
  function calibrar(tipo) { return calibrarAgora(tipo); }
  async function calibrarAgora(tipo) {
    if (!S.dev || S.ocupado) return;
    S.ocupado = true;
    status(`Calibrando ${tipo}…`);
    try {
      const ok = await S.dev.calibrar(tipo);
      S.cal = await S.dev.lerCalibracao();
      S.ocupado = false;
      S.aviso = ok ? '' : `O aparelho recusou a calibração do ${tipo}. Tente de novo.`;
      status(`Calibração ${tipo}: ${ok ? 'concluída' : 'FALHOU'}`);
    } catch (e) { erroDoAparelho(e); }
  }

  function erroDoAparelho(e) {
    S.ocupado = false;
    log('erro: ' + e.message);
    S.aviso = /incompleta/.test(e.message)
      ? 'O aparelho não respondeu inteiro — pode ter entrado em repouso. Aperte o botão dele e tente de novo.'
      : e.message;
    status('Pronto');
  }

  async function medir(tipo) {
    if (!S.dev || S.ocupado) return;
    if (S.os && S.av && S.av.n == null) { S.aviso = avDicaFase(); desenhar(); return; }
    S.ocupado = true;
    status('Medindo…');
    try { chegou(await S.dev.medir(modoAtual()), tipo); } catch (e) { erroDoAparelho(e); }
  }

  // O botão do próprio DS-01 foi apertado: a medição já foi feita, só baixar.
  /* O BOTÃO DO APARELHO RESPEITA O MODO ESCOLHIDO (30/10/2026, Jones: "faça essas
     modificações no programa"). O DS-01 mede pelo botão sempre no modo dele
     (SCI); o programa só baixava essa leitura, e o "Modo de medição" das
     Configurações valia só no F5/F6. Com SCE escolhido, 6 de 7 medições saíram
     em SCI sem ninguém ver. Agora, se o modo do botão não é o escolhido, o
     programa mede de novo na hora, no modo certo — o aparelho ainda está
     encostado na peça. */
  function aoBotao(modo) {
    if (S.ocupado || !S.dev) return;
    let tipo = cfg.botao;
    if (tipo === 'amostra' && !S.padrao && !S.acumulando.length) tipo = 'padrao';
    /* O BOTÃO DO APARELHO NA AVALIAÇÃO DE OS (12/11/2026, Junior: "como as
       comparações são automáticas com base nas medições de referência, o botão
       deve ter função de medir amostra como referência para todas as amostras").
       Aqui não há padrão nem escolha: o botão mede SEMPRE a amostra atual, e a
       média dela passa a valer como referência para todas as outras da fase (a
       tabela de compatibilidade é refeita na hora). O "mede como Amostra/Padrão"
       das Configurações só vale fora da OS. Sem amostra atual (total já medido),
       o aperto não some em silêncio: a tela diz o que fazer. */
    if (S.os && S.av) {
      if (S.av.n == null) { S.aviso = 'Botão do aparelho: ' + avDicaFase(); desenhar(); return; }
      tipo = 'amostra';
    }
    S.ocupado = true;
    const alvo = modoAtual();
    if (cfg.modo !== 'auto' && modo !== alvo) {
      log(`botão do aparelho mediu em ${MODOS[modo] || modo}; medindo de novo em ${MODOS[alvo] || alvo}`);
      status(`Botão do aparelho mediu em ${MODOS[modo] || modo} — medindo de novo em ${MODOS[alvo] || alvo}… não tire o aparelho da peça`);
      S.dev.medir(alvo).then(m => chegou(m, tipo), erroDoAparelho);
      return;
    }
    status('Medição feita no aparelho — baixando…');
    S.dev.baixarMedicao(modo).then(m => chegou(m, tipo), erroDoAparelho);
  }

  function chegou(m, tipo) {
    S.ocupado = false;
    const n = Math.max(1, parseInt(cfg.media, 10) || 1);
    if (S.acumulando.length && S.acumulando[0][1] !== tipo) S.acumulando = [];
    S.acumulando.push([m, tipo]);
    if (S.acumulando.length < n) {
      status(`Média: leitura ${S.acumulando.length} de ${n} registrada — reposicione e meça de novo`);
      return;
    }
    const medidas = S.acumulando.map(x => x[0]);
    S.acumulando = [];
    let esp = [], labAp = null;
    if (medidas[0].espectro.length) {
      esp = medidas[0].espectro.map((_, i) => medidas.reduce((s, x) => s + (x.espectro[i] || 0), 0) / medidas.length);
    } else {
      labAp = [0, 1, 2].map(i => medidas.reduce((s, x) => s + x.lab_aparelho[i], 0) / medidas.length);
    }
    const naOS = !!(S.os && S.av);
    if (!naOS) S.cont[tipo]++;
    const prefixo = tipo === 'padrao' ? cfg.nome_padrao : cfg.nome_amostra;
    const l = novaLeitura({ nome: `${prefixo} ${S.cont[tipo]}`, tipo, espectro: esp, wave_start: medidas[0].wave_start,
      lab_aparelho: labAp, horario: new Date(), modo: medidas[0].modo });
    // Com a tela ligada a uma OS, quem dá nome, referência e tom é a avaliação.
    const msg = naOS ? avMedicao(l) : '';
    if (!naOS && tipo === 'amostra' && S.padrao) l.refId = S.padrao.id;
    S.leituras.unshift(l);
    if (tipo === 'padrao' && !naOS) S.padrao = l;
    S.atual = l;
    guardar(l);
    status(msg || `${l.nome} medido às ${horaTxt(l.horario)}`);
  }

  // Cada medição vai para o histórico deste navegador na hora, como o CSV do dia do original.
  function registroDe(l) {
    const ref = refDe(l);
    return { id: l.id, dia: diaIso(l.horario), horario: l.horario.toISOString(), nome: l.nome, tipo: l.tipo,
      espectro: l.espectro, wave_start: l.wave_start, lab_aparelho: l.lab_aparelho,
      refId: l.refId, refNome: ref ? ref.nome : '', iluminante: cfg.iluminante1, modo: l.modo,
      osId: l.osId || null, osNumero: l.osNumero || '', amostraN: l.amostraN || null, rodada: l.rodada || null,
      // A fase da OS (12/11/2026): sem ela, a medição só podia ser posta de volta
      // na avaliação certa adivinhando pela ordem.
      fase: l.fase != null ? l.fase : null, faseNome: l.faseNome || '' };
  }
  function guardar(l) {
    BD.gravar(registroDe(l)).catch(e => {
      S.semBD = true;
      log('não consegui guardar no histórico: ' + (e && e.message));
      desenhar();
    });
  }

  /* ---------------- lista da sessão ---------------- */
  function selecionar(id) {
    const l = S.leituras.find(x => x.id === id);
    if (l) { S.atual = l; desenhar(); }
  }
  function usarPadrao(id) {
    const l = S.leituras.find(x => x.id === id);
    if (!l) return;
    l.tipo = 'padrao'; l.refId = null;
    S.padrao = l; S.atual = l;
    guardar(l);
    desenhar();
  }
  function renomear(id) {
    const l = S.leituras.find(x => x.id === id);
    if (!l) return;
    perguntar({ titulo: 'Renomear', texto: `Novo nome para "${l.nome}":`, campo: l.nome, sim: 'Renomear',
      executar: novo => { if (novo && novo.trim()) { l.nome = novo.trim(); guardar(l); } desenhar(); } });
  }
  // Como no original: sai da lista da tela; o histórico do dia guarda a medição.
  function excluir(id) {
    const i = S.leituras.findIndex(x => x.id === id);
    if (i < 0) return;
    const [l] = S.leituras.splice(i, 1);
    if (l === S.padrao) S.padrao = null;
    if (l === S.atual) S.atual = null;
    desenhar();
  }
  function limpar() {
    if (!S.leituras.length) return;
    perguntar({ titulo: 'Limpar a lista', texto: 'Apagar a lista da tela? O histórico do dia continua guardado.', sim: 'Limpar', lugar: 'lista',
      executar: () => {
        S.leituras = []; S.atual = null; S.cont = { padrao: 0, amostra: 0 };
        // Na avaliação de uma OS a referência do tom é da avaliação, não da
        // lista: continua valendo depois de limpar.
        S.padrao = S.os && S.av && S.av.n != null && estaNumTom(S.av.n) ? refComoLeitura(tomDe(S.av.n)) : null;
        desenhar();
      } });
  }

  /* ---------------- CSV (o mesmo formato do original) ---------------- */
  const cabecalhoCsv = () => ['Data', 'Hora', 'Nome', 'Tipo', 'Iluminante', 'L*', 'a*', 'b*', 'C*', 'h°',
    'Padrão de referência', 'dL*', 'da*', 'db*', ...FORMULAS, 'Resultado']
    .concat(Array.from({ length: 31 }, (_, i) => `${400 + 10 * i} nm`));
  function linhaCsv(l, ref) {
    const [C, h] = lch(...l.lab1);
    const linha = [dataTxt(l.horario), horaTxt(l.horario), l.nome, l.tipo === 'padrao' ? 'Padrão' : 'Amostra', cfg.iluminante1,
      ...[...l.lab1, C, h].map(v => num(v, 3))];
    if (l.tipo === 'amostra' && ref && ref !== l) {
      const r = comparar(ref, l);
      linha.push(ref.nome, ...r.d.slice(0, 3).map(v => num(v, 3)), ...FORMULAS.map(f => num(r.des[f], 3)), r.aprovado ? 'Aprovado' : 'Reprovado');
    } else {
      linha.push(...new Array(9).fill(''));
    }
    const esp = {};
    (l.espectro || []).forEach((v, i) => { esp[l.wave_start + 10 * i] = v; });
    for (let w = 400; w <= 700; w += 10) linha.push(w in esp ? num(esp[w], 3) : '');
    return linha;
  }
  function baixarCsv(nome, linhas) {
    const q = c => /[;"\n]/.test(c) ? '"' + String(c).replace(/"/g, '""') + '"' : c;
    const txt = '\uFEFF' + linhas.map(l => l.map(q).join(';')).join('\r\n') + '\r\n';
    const url = URL.createObjectURL(new Blob([txt], { type: 'text/csv;charset=utf-8' }));
    const a = document.createElement('a');
    a.href = url; a.download = nome;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }
  /* EXPORTAR A AVALIAÇÃO DA OS (12/11/2026). O "Exportar CSV" da sessão só leva
     a lista da tela, e o do Histórico só um dia daquele navegador; a avaliação
     da OS guarda TODAS as medições de todas as fases, de qualquer dia. Para
     conferir a compatibilidade fora do programa (caso OS 0715), este CSV traz
     cada medição: fase, amostra, horário, L*a*b*, a média da amostra e o tom em
     que ela está agora. Só leitura — não muda nada da avaliação. */
  async function exportarAvaliacao() {
    if (!S.avOS || !S.os || !S.fases) { S.aviso = 'Escolha a OS antes de exportar a avaliação.'; desenhar(); return; }
    sincronizarRibanas();
    const v = x => (x == null ? '' : String(Math.round(x * 1000) / 1000).replace('.', ','));
    const linhas = [];
    const ids = new Set();
    // a) o que está na avaliação guardada neste navegador
    S.fases.forEach(f => {
      const av = S.avOS.fases[f.i];
      if (!av) return;
      numerosDasAmostras(av).forEach(n => {
        const media = corDaAmostra(n, av) || [];
        const t = tomDe(n, av);
        const tomTxt = t === 0 ? 'não usar' : t === 'x' ? 'incompatível com o corpo' : t == null ? '' : 'Tom ' + t;
        const fixo = av.fixo && (n in av.fixo) ? (av.fixo[n] === 0 ? 'não usar' : 'Tom ' + av.fixo[n]) : '';
        (av.medidas[n] || []).forEach((m, j) => {
          if (m.id) ids.add(m.id);
          const em = m.em ? new Date(m.em) : null;
          linhas.push({ t: em ? em.getTime() : 0, c: [S.os.numero, f.nome, f.tecido || '', n, j + 1, em ? `${dataTxt(em)} ${horaTxt(em)}` : '',
            v(m.lab && m.lab[0]), v(m.lab && m.lab[1]), v(m.lab && m.lab[2]), v(media[0]), v(media[1]), v(media[2]), tomTxt, fixo, 'avaliação'] });
        });
      });
    });
    // b) o histórico deste navegador: toda medição desta OS, de qualquer dia,
    //    que não esteja na avaliação (ex.: avaliação perdida no F5). A fase só é
    //    conhecida nas medições gravadas a partir de 12/11/2026.
    let doHist = 0;
    try {
      const regs = (await BD.todos()).filter(r => r && r.osId === S.os.id && !ids.has(r.id));
      regs.forEach(r => {
        const l = leituraDoRegistro(r);
        const f = r.fase != null ? S.fases[r.fase] : null;
        linhas.push({ t: l.horario.getTime(), c: [S.os.numero, f ? f.nome : (r.faseNome || 'não informada (histórico)'), f ? (f.tecido || '') : '',
          r.amostraN || '', '', `${dataTxt(l.horario)} ${horaTxt(l.horario)}`,
          v(l.lab1 && l.lab1[0]), v(l.lab1 && l.lab1[1]), v(l.lab1 && l.lab1[2]), '', '', '', '', '', 'histórico'] });
        doHist++;
      });
    } catch (e) { log('exportar avaliação: histórico ilegível: ' + (e && e.message)); }
    if (!linhas.length) { S.aviso = 'Não há medições desta OS neste navegador (nem na avaliação, nem no Histórico).'; desenhar(); return; }
    linhas.sort((p, q) => p.t - q.t);
    baixarCsv(`avaliacao_${S.os.doc || "OS"}${S.os.numero}_${diaIso(new Date())}.csv`,
      [['OS', 'Fase', 'Tecido', 'Amostra', 'Medição', 'Data/hora', 'L*', 'a*', 'b*', 'Média L*', 'Média a*', 'Média b*', 'Tom agora', 'Fixado à mão', 'Origem'], ...linhas.map(x => x.c)]);
    status(`${S.os.nome || "OS " + S.os.numero} exportada: ${linhas.length - doHist} medições da avaliação + ${doHist} só do Histórico deste navegador.`);
  }

  function exportar() {
    if (!S.leituras.length) { S.aviso = 'Nenhuma medição na lista.'; desenhar(); return; }
    const agora = new Date();
    const lista = S.leituras.slice().reverse();   // na ordem em que foram medidas
    baixarCsv(`medicoes_${diaIso(agora)}_${dois(agora.getHours())}${dois(agora.getMinutes())}.csv`,
      [cabecalhoCsv(), ...lista.map(l => linhaCsv(l, refDe(l)))]);
  }

  // Lê um CSV do original (ou exportado daqui) e põe no histórico deste navegador.
  async function importar(arquivo) {
    if (!arquivo) return;
    try {
      await carregarTabelas();
      let txt = await arquivo.text();
      txt = txt.replace(/^\uFEFF/, '');
      const linhas = txt.split(/\r?\n/).filter(x => x.trim()).map(x => x.split(';'));
      const cab = linhas[0] || [];
      const col = {};
      cab.forEach((n, i) => { col[n] = i; });
      if (!('L*' in col) || !('Data' in col)) throw new Error('não parece um CSV do colorímetro');
      const ondas = cab.map((n, i) => [i, n]).filter(([, n]) => / nm$/.test(n)).map(([i, n]) => [i, parseInt(n, 10)]);
      const padroes = {};
      let n = 0;
      for (const ln of linhas.slice(1)) {
        const fl = k => lerNum(ln[col[k]] || '');
        const labL = ['L*', 'a*', 'b*'].map(fl);
        const [d, mth, a] = String(ln[col.Data] || '').split('/').map(Number);
        const [hh, mm, ss] = String(ln[col.Hora] || '').split(':').map(Number);
        if (labL.some(v => v == null) || !d || !a) continue;
        const horario = new Date(a, mth - 1, d, hh || 0, mm || 0, ss || 0);
        const esp = ondas.map(([i]) => lerNum(ln[i] || ''));
        const espectro = esp.length && esp.every(v => v != null) ? esp : [];
        const tipo = ln[col.Tipo] === 'Padrão' ? 'padrao' : 'amostra';
        const reg = { id: 'imp_' + horario.getTime() + '_' + (++seq), dia: diaIso(horario), horario: horario.toISOString(),
          nome: ln[col.Nome] || '', tipo, espectro, wave_start: ondas.length ? ondas[0][1] : 400,
          lab_aparelho: labL, refId: null, refNome: '', iluminante: ln[col.Iluminante] || cfg.iluminante1, importado: arquivo.name };
        // O contador do original recomeça a cada abertura: vale o último padrão com esse nome.
        if (tipo === 'amostra') {
          const p = padroes[ln[col['Padrão de referência']]];
          if (p) { reg.refId = p.id; reg.refNome = p.nome; }
        } else padroes[reg.nome] = reg;
        await BD.gravar(reg);
        n++;
      }
      S.aviso = `${n} medições de "${arquivo.name}" entraram no histórico.`;
      await abrirHistorico(true);
    } catch (e) {
      S.aviso = 'Não consegui importar: ' + e.message;
      desenhar();
    }
  }

  /* ---------------- histórico ---------------- */
  function leituraDoRegistro(r) {
    const l = novaLeitura({ id: r.id, nome: r.nome, tipo: r.tipo, espectro: r.espectro || [], wave_start: r.wave_start || 400,
      lab_aparelho: r.lab_aparelho, horario: new Date(r.horario), modo: r.modo });
    l.refId = r.refId; l.refNome = r.refNome; l.osId = r.osId || null; l.osNumero = r.osNumero || '';
    return l;
  }
  async function abrirHistorico(manter) {
    // O quadro troca NA HORA do clique ("Carregando…"); esperar a leitura para
    // trocar deixava o quadro anterior na tela, e o clique seguinte agia nele.
    S.janela = 'hist';
    S.histCarregando = true;
    desenhar();
    try {
      await carregarTabelas();
      const todos = await BD.todos();
      S.histDias = [...new Set(todos.map(r => r.dia))].sort().reverse();
      if (!manter || !S.histDias.includes(S.histDia)) S.histDia = S.histDias[0] || '';
      S.histCarregando = false;
      if (S.janela !== 'hist') return;   // fechado enquanto carregava
      await carregarDia(S.histDia);
    } catch (e) {
      S.histCarregando = false;
      S.aviso = 'O histórico deste navegador não abriu: ' + (e && e.message);
      desenhar();
    }
  }
  async function carregarDia(dia) {
    S.histDia = dia;
    const regs = dia ? await BD.doDia(dia) : [];
    regs.sort((a, b) => String(a.horario).localeCompare(String(b.horario)));
    S.histLidas = regs.map(leituraDoRegistro);
    desenhar();
  }
  function histUsarPadrao(id) {
    const r = S.histLidas.find(x => x.id === id);
    if (!r) return;
    const l = novaLeitura({ nome: r.nome, tipo: 'padrao', espectro: r.espectro, wave_start: r.wave_start,
      lab_aparelho: r.lab_aparelho, horario: r.horario, modo: r.modo });
    S.leituras.unshift(l);
    S.padrao = l; S.atual = l;
    S.janela = null;
    status(`${r.nome} de ${dataTxt(r.horario).slice(0, 5)} ${horaTxt(r.horario).slice(0, 5)} virou o padrão. Agora meça a amostra`);
  }
  function histExportar() {
    if (!S.histLidas.length) return;
    const porId = {};
    S.histLidas.forEach(l => { porId[l.id] = l; });
    baixarCsv(`medicoes_${S.histDia}.csv`, [cabecalhoCsv(), ...S.histLidas.map(l => linhaCsv(l, porId[l.refId]))]);
  }

  /* UMA JANELA DE CADA VEZ (19/10/2026, Junior: "as janelas de configuração e
     histórico são abertas, mas não podem ser fechadas. Elas se acumulam").
     Configurações e Histórico dividem o mesmo lugar no topo da tela: abrir uma
     fecha a outra, o botão que abriu fecha de novo, o ✕ do título fecha e o Esc
     também. Antes cada uma tinha o seu "mostrar", abrir o Histórico deixava as
     Configurações abertas, e o único Fechar ficava no fim do quadro. */
  function botaoFechar() {
    return '<button type="button" class="btn small ghost cm-fechar" title="Fechar (Esc)" onclick="Colorimetro.fecharJanela()">✕ Fechar</button>';
  }
  function fecharJanela() { S.janela = null; desenhar(); }
  function alternar(qual, abrir) {
    if (S.janela === qual && !abrir) { fecharJanela(); return; }
    S.pedido = null;
    if (qual === 'hist') { abrirHistorico(false); rolarParaJanela(); return; }
    if (qual === 'rel') { abrirRelatorio(); return; }
    S.janela = qual;
    desenhar();
    rolarParaJanela();
  }
  function rolarParaJanela() {
    const j = document.querySelector('#colorimetro-painel .cm-cfg, #colorimetro-painel .cm-hist, #colorimetro-painel .cm-rel');
    if (j && j.scrollIntoView) j.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  /* ---------------- AVALIAR AMOSTRAS DE UMA OS ----------------
     (19/10/2026, Junior, em três recados:
      1. "As medições devem apontar para uma OS. O usuário deve clicar numa
         opção na coluna de ações da lista de OS 'avaliar amostras' e ser
         redirecionado para a área Colorímetro… relatório conduzido pelo usuário
         para determinar quais cores são compatíveis em conjunto numérico, onde
         mais de um grupo pode ser formado" — e "esse relatório deve constar no
         campo de observações da OS escolhida".
      2. "deve obedecer à sequência de amostras, iniciando na amostra número 1 e
         passando pelas amostras número 2, 3, 4, 5…".
      3. "deve obedecer à mudança de uma amostra para outra, pois pode exigir
         mais de uma medição por amostra. Usuário deve clicar para mudar o número
         da amostra que está sendo avaliada. Essa amostra deve migrar para grupo
         de tonalidades compatíveis automaticamente conforme receba pelo menos uma
         aprovação. Quando usuário clicar em trocar, significa que aquela amostra
         deve ir para um grupo que deverá ser novamente avaliado para encontrar o
         grupo de compatibilidade de tonalidade.")

      4. "Em casos de OS com mais de uma fase de enfesto, o programa deve gerar o
         relatório de grupo numérico de amostras compatíveis de acordo com as
         fases de enfesto e tipo de tecido. Logo, usuário deve seguir as fases de
         enfesto na hora de avaliar as amostras. Caso ele tenha que repetir a
         avaliação de amostras, ele deve apontar qual fase está sendo
         reavaliada.")

     POR FASE DE ENFESTO: cada fase da OS (Corpo · Malha Algodão, Gola · Ribana…)
     tem a sua avaliação inteira — amostras numeradas desde 1, rodadas e tons
     próprios. A tela anda na ORDEM das fases: só se mede a fase atual, e
     "Concluir a fase" passa para a seguinte. Repetir uma fase já concluída é
     escolhê-la em "Reavaliar a fase" — a avaliação dela recomeça do zero e o
     relatório diz que ela foi reavaliada.

     A DINÂMICA de cada fase, então, é por RODADAS, e cada rodada forma um tom:
       · A AMOSTRA ATUAL é um número (1, 2, 3…). Medir mede ESSA amostra, quantas
         vezes for preciso; o número só muda no clique (Trocar / Ir para nº).
       · A REFERÊNCIA DO TOM é sempre uma AMOSTRA: a primeira medição da rodada,
         ou a amostra atual medida como referência (avReferencia) — que já entra
         no tom, porque é ela que o define. Na 1ª rodada é a Amostra 1.
       · Cada medição da amostra atual é comparada com a referência (fórmula e
         tolerância das Configurações). UMA aprovação basta: a amostra entra no
         tom da rodada na hora, sozinha. Medições reprovadas depois não a tiram.
       · TROCAR sem aprovação manda a amostra para "a reavaliar".
       · REAVALIAR abre a rodada seguinte (Tom 2, Tom 3…) só com as que estão a
         reavaliar, na ordem dos números; a primeira delas vira a referência do
         tom novo. Cada rodada tira ao menos uma da fila, então ela termina.
     O relatório mostra o resultado, deixa a pessoa corrigir o tom de qualquer
     amostra e grava o resumo como recado nas observações da OS (salvarObsNota:
     assinado, com data, sem passar por cima de ninguém).

     O estado mora neste navegador, por OS (localStorage), junto com as medições
     (IndexedDB): fechar a tela e voltar à OS continua de onde parou. */
  const CHAVE_AV = id => 'gos:colorimetro:av:' + id;
  function rotuloOS(o) {
    const g = o.grade && o.grade.descricao ? ' · ' + o.grade.descricao : '';
    return `OS ${o.os || '—'}${o.modeloNome ? ' · ' + o.modeloNome : ''}${g}`;
  }
  /* A ORDEM DE PREPARO DE MATÉRIA-PRIMA — OP (12/11/2026, Junior: "Avaliar
     amostras, da mesma forma como em uma OS"). Um lote de bobinas de tecido e
     de ribana do Estoque de tecidos (STATE.ordensPreparo). A tela é a mesma da
     OS: S.os guarda qual documento é (doc 'OS' ou 'OP') e o nome dele. As
     fases: Tecido (uma amostra por bobina) e Ribana, comparada com o Tecido. */
  const listaOS = () => (typeof STATE !== 'undefined' && STATE && STATE.ordens) || [];
  const listaOP = () => (typeof STATE !== 'undefined' && STATE && STATE.ordensPreparo) || [];
  function acharDoc(id) {
    const o = listaOS().find(x => x.id === id);
    if (o) return { tipo: 'OS', o };
    const p = listaOP().find(x => x.id === id);
    return p ? { tipo: 'OP', o: p } : null;
  }
  const opNum = p => String(p.numero || 0).padStart(4, '0');
  const rotuloOP = p => `OP ${opNum(p)}${[p.tecidoNome, p.corNome].filter(Boolean).length ? ' · ' + [p.tecidoNome, p.corNome].filter(Boolean).join(' ') : ''}`;
  function fasesDaOP(p) {
    const cor = p.corNome || '', bt = Number(p.bobinasTecido) || 0, br = Number(p.bobinasRibana) || 0;
    const f = [];
    if (bt > 0 || !br) f.push({ nome: 'Tecido', tecido: [p.tecidoNome, cor].filter(Boolean).join(' '), total: bt });
    if (br > 0) f.push({ nome: 'Ribana', tecido: [p.ribanaNome || 'Ribana', cor].filter(Boolean).join(' '), total: br, ribana: true });
    return f.map((x, i) => ({ i, nome: x.nome, tecido: x.tecido, total: x.total, semAval: false,
      corDe: x.ribana && i > 0 ? 0 : -1, rotulo: `Fase ${i + 1} · ${x.nome}${x.tecido ? ' · ' + x.tecido : ''}` }));
  }
  function avNovo(osId) { return { osId, n: 1, rodada: 1, tons: {}, ref: {}, medidas: {}, fila: null, fimRodada: false, fixo: {}, modelo: 'dados' }; }
  // Por OS: a fase atual, a avaliação de cada fase, as concluídas e quantas vezes cada uma foi reavaliada.
  function avNovoOS(osId) { return { osId, versao: 2, faseAtual: 0, fases: {}, concluidas: {}, reav: {} }; }
  function avCarregar(osId) {
    try {
      const v = JSON.parse(localStorage.getItem(CHAVE_AV(osId)) || 'null');
      if (v && v.osId === osId && v.versao === 2) return Object.assign(avNovoOS(osId), v);
      if (v && v.osId === osId && v.medidas) { const c = avNovoOS(osId); c.fases[0] = Object.assign(avNovo(osId), v); return c; }
    } catch (e) { /* sem storage */ }
    return avNovoOS(osId);
  }
  function avGuardar() {
    if (!S.avOS) return;
    try { localStorage.setItem(CHAVE_AV(S.avOS.osId), JSON.stringify(S.avOS)); S.avNaoGuardou = false; }
    catch (e) {
      // ANTES ERA CALADO (12/11/2026, caso OS 0715): sem espaço no navegador, a
      // avaliação seguia só na memória e sumia no F5 — o corpo da 0715 voltou com
      // 3 medições. As medições seguem no histórico (IndexedDB); o aviso diz isso.
      S.avNaoGuardou = true;
      S.aviso = 'Este navegador não conseguiu guardar a avaliação da OS (' + ((e && e.name) || 'erro') + '). '
        + 'As medições continuam no Histórico; use "Exportar dados da avaliação" antes de recarregar a página.';
    }
  }
  // As fases de enfesto da OS, com o tecido e a cor de cada uma.
  function fasesDaOS(osId) {
    const d = acharDoc(osId);
    if (d && d.tipo === 'OP') return fasesDaOP(d.o);
    const o = (typeof STATE !== 'undefined' && STATE && STATE.ordens || []).find(x => x.id === osId);
    const lista = o && Array.isArray(o.fases) && o.fases.length ? o.fases : [{ nome: 'Fase única' }];
    return lista.map((f, i) => {
      let cor = f.corNome || (o && o.enfesto && o.enfesto.blocos && o.enfesto.blocos[i] && o.enfesto.blocos[i].nomeCor)
        || (o && o.tecidos && o.tecidos[i] && o.tecidos[i].corNome) || '';
      if (cor === '—') cor = '';
      try { if (cor && typeof corSemTecido === 'function') cor = corSemTecido(cor, f.tecidoNome) || cor; } catch (e) { /* fica como está */ }
      // "Vermelho Malha Algodão" com o tecido "Malha Algodão": a cor é só "Vermelho".
      const tec = String(f.tecidoNome || '').trim();
      if (cor && tec && cor.toLowerCase().endsWith(tec.toLowerCase())) cor = cor.slice(0, cor.length - tec.length).trim() || cor;
      const tecido = [f.tecidoNome, cor].filter(Boolean).join(' ');
      return { i, nome: f.nome || `Fase ${i + 1}`, tecido, rotulo: `Fase ${i + 1} · ${f.nome || 'sem nome'}${tecido ? ' · ' + tecido : ''}`,
        corDe: ehRibana(f) ? faseDeCorDaRibana(o, lista) : -1, semAval: faseSemAvaliacao(o, f) };
    });
  }
  /* FASE QUE NÃO PRECISA DE AVALIAÇÃO (20/10/2026, Junior: "cm.rec fase corpo 2
     não necessita ter tecidos compatíveis em cor, pois não há outras peças do
     produto que são unidas para que as cores sejam compatíveis com branco").
     Avaliar amostras serve para casar a cor de peças costuradas juntas; a fase
     cujas peças não se unem a nenhuma outra da mesma cor fica fora: aparece na
     lista como "não precisa avaliar", a tela pula por ela e o relatório não a
     cita. Os índices das fases continuam os da OS (a avaliação guardada é por
     índice), então nada que já foi medido muda de lugar.
     Depois: "CM.TRI apenas a fase corpo 1 precisa receber avaliação de
     amostras" — Corpo 2 e Corpo 3 saem; ribana e viés seguem como estão. */
  const CORPO_N = n => new RegExp(`^corpo\\s*(parte\\s*)?${n}\\b`, 'i');
  const FASES_SEM_AVALIACAO = {
    'CM.REC': [CORPO_N(2)],
    'CM.TRI': [CORPO_N(2), CORPO_N(3)], 'CM.TRICOLOR': [CORPO_N(2), CORPO_N(3)]
  };
  // Viés não precisa avaliar em linha nenhuma (20/10/2026, Junior: "viés não precisa avaliar").
  const VIES_PURO = /^vi[eé]s\b/i;
  // BM.TRI avalia TODAS as fases, viés inclusive (20/10/2026, Junior: "bm.t todas as fases precisam ser avaliadas").
  const LINHAS_AVALIA_TUDO = ['BM.TRI', 'BM.TRICOLOR'];
  function faseSemAvaliacao(o, f) {
    const linha = ((o && o.grade && o.grade.descricao || '').toUpperCase().match(/\b(CM|BM)\.[A-Z]+/) || [''])[0];
    if (LINHAS_AVALIA_TUDO.includes(linha)) return false;
    if (VIES_PURO.test(String(f.nome || '').trim())) return true;
    return (FASES_SEM_AVALIACAO[linha] || []).some(re => re.test(String(f.nome || '').trim()));
  }
  // Fase que a tela não precisa mais medir: concluída ou sem avaliação.
  const faseFeita = (c, j) => !!(c.concluidas[j] || (S.fases && S.fases[j] && S.fases[j].semAval));
  /* A RIBANA VAI JUNTO DA FASE DE COR DELA (20/10/2026, Junior: "A fase ribana
     deve sempre ser conjugada no relatório de acordo com a fase de cor
     correspondente. Quando cm.lisa, ribana igual a corpo, quando cm.rec, ribana
     igual a corpo 1, quando bm.lisa, ribana igual a corpo, quando bm.tri ribana
     igual a corpo 3"; e depois "Na CM.TRI, ribana igual a corpo 1"). A ribana
     é tinta na cor de UMA das fases de corpo; no relatório ela não sai sozinha,
     sai dentro dessa fase, tom a tom (o Tom 1 da ribana com o Tom 1 do corpo).
     Linha fora da lista e com um corpo só (CM.B, BM.B…): esse corpo; mais de um
     corpo e linha fora da lista: a ribana fica separada, como antes. */
  const ehRibana = f => /ribana/i.test(String(f.tecidoNome || '') + ' ' + String(f.nome || ''));
  const CORPO_DA_RIBANA = { 'CM.LISA': 0, 'BM.LISA': 0, 'CM.REC': 1, 'CM.TRI': 1, 'CM.TRICOLOR': 1, 'BM.TRI': 3, 'BM.TRICOLOR': 3 };
  function faseDeCorDaRibana(o, lista) {
    const linha = ((o && o.grade && o.grade.descricao || '').toUpperCase().match(/\b(CM|BM)\.[A-Z]+/) || [''])[0];
    const corpos = lista.map((f, i) => ({ i, nome: String(f.nome || '') })).filter(x => /^corpo/i.test(x.nome) && !ehRibana(lista[x.i]));
    const p = CORPO_DA_RIBANA[linha];
    if (p === 0 || (p == null && corpos.length === 1)) return corpos.length ? corpos[0].i : -1;
    if (p == null) return -1;
    const c = corpos.find(x => new RegExp(`^corpo\\s*(parte\\s*)?${p}\\b`, 'i').test(x.nome));
    return c ? c.i : -1;
  }
  // Põe a fase i na tela: S.av passa a ser a avaliação dela.
  function avUsarFase(i) {
    const c = S.avOS;
    c.faseAtual = i;
    if (!c.fases[i]) c.fases[i] = avNovo(c.osId);
    S.av = c.fases[i];
    avNormalizar(S.av);
    classificar(S.av);
    S.padrao = S.av.n != null && estaNumTom(S.av.n) ? refComoLeitura(tomDe(S.av.n)) : null;
    S.atual = null;
  }

  const deDaFormula = (a, b) => {
    const f = cfg.formula;
    if (f === 'dE00') return de2000(a, b, ...cfg.k00);
    if (f === 'dEcmc') return deCmc(a, b, cfg.cmc_l, cfg.cmc_c);
    if (f === 'dE94') return de94(a, b, ...cfg.k94);
    return deAb(a, b);
  };
  const tolAtual = () => cfg.tol[cfg.formula];
  const numerosDasAmostras = (av = S.av) => Object.keys(av.medidas).map(Number).filter(n => n > 0).sort((a, b) => a - b);
  const tomDe = (n, av = S.av) => av.tons[n];
  const estaNumTom = (n, av = S.av) => typeof tomDe(n, av) === 'number' && tomDe(n, av) > 0;

  /* CLASSIFICAÇÃO PELOS DADOS MEMORIZADOS, SEM REPESCAGEM (12/11/2026, Junior:
     "a plataforma deve classificar quais amostras são compatíveis sem necessitar
     de repescagem, mas sim apenas com os dados coletados no momento da medição
     devem ser suficientes para indicar compatibilidade de cores com os dados
     memorizados das amostras").
     Antes a fase andava por RODADAS: o tom se formava contra uma referência, a
     reprovada ia para "a reavaliar" e era preciso medir tudo de novo na rodada do
     Tom 2 (repescagem). Agora cada amostra é medida uma vez (ou mais, se quiser)
     e a cor dela fica guardada; a cada medição o programa refaz a classificação
     da fase inteira com o que já tem:
       · a cor da amostra = a média de todas as medições dela;
       · duas amostras são compatíveis quando passam na mesma regra de aprovação
         da tela (dE escolhido E dL*, da*, db* dentro das tolerâncias — comparar);
       · as amostras entram em tons na ordem do número: cada uma vai para o tom
         em que é compatível com TODAS as que já estão nele (o de menor dE, se
         couber em mais de um); sem nenhum, abre o tom seguinte. Assim duas
         peças do mesmo tom sempre casam entre si, não só com a referência;
       · o que se marcou à mão no relatório (av.fixo: um tom, ou 0 = não usar)
         vale por cima e entra primeiro.
     A referência de cada tom passa a ser só a primeira amostra dele (a de menor
     número), para a tela mostrar a comparação.
     SUBSTITUÍDO NO MESMO DIA pelo ROLO DE REFERÊNCIA (opção A, ver dentro de
     classificar): a média e o "todas com todas" não são o que a conferência a
     olho faz. A média segue só como a cor exibida da amostra. */
  const corDaAmostra = (n, av = S.av) => {
    const ms = (av.medidas[n] || []).filter(m => m && m.lab);
    return ms.length ? [0, 1, 2].map(i => ms.reduce((s, m) => s + m.lab[i], 0) / ms.length) : null;
  };
  const compativeis = (a, b) => comparar({ lab1: a }, { lab1: b }).aprovado;
  // A leitura do rolo de referência: a 1ª medição dele, como a referência de antes.
  const labRef = (refN, av = S.av) => { const m = (av.medidas[refN] || []).find(x => x && x.lab); return m ? m.lab : null; };
  // Amostra n contra o rolo de referência refN: alguma medição dela aprovada?
  function aprovadaContraRef(n, refN, av = S.av) {
    const m = melhorContra(av.medidas[n], labRef(refN, av));
    return !!(m && m.aprovado);
  }
  // A melhor de várias medições contra a leitura de uma referência (aprovada antes
  // de reprovada; entre iguais, a de menor dE). Serve também à ribana × corpo.
  function melhorContra(medidas, refLab) {
    if (!refLab) return null;
    let melhor = null;
    (medidas || []).forEach(m => {
      if (!m || !m.lab) return;
      const r = comparar({ lab1: refLab }, { lab1: m.lab });
      const d = r.des[cfg.formula];
      if (!melhor || (r.aprovado && !melhor.aprovado) || (r.aprovado === melhor.aprovado && d < melhor.d)) melhor = { r, d, aprovado: r.aprovado };
    });
    return melhor;
  }
  // A melhor medição de n contra a referência: { r (comparar), aprovado } — para a tela.
  function melhorContraRef(n, refN, av = S.av) {
    return melhorContra(av.medidas[n], labRef(refN, av));
  }
  function classificar(av = S.av) {
    if (!av) return;
    av.fixo = av.fixo || {};
    // Ribana com corpo medido: classificada CONTRA o corpo (ver classificarRibana).
    const ctxCorpo = corpoDaRibana(av);
    if (ctxCorpo) { classificarRibana(av, ctxCorpo); return; }
    av.contraCorpo = false;
    const cores = {};
    numerosDasAmostras(av).forEach(n => { const c = corDaAmostra(n, av); if (c) cores[n] = c; });
    const nums = Object.keys(cores).map(Number).sort((a, b) => a - b);
    const tons = {};                               // k -> [n]
    const tonsNovos = {};
    // "Só esta amostra" que ninguém mais fixa deixa de valer (a amostra voltou ao automático).
    if (av.tomSo) Object.keys(av.tomSo).forEach(k => { if (!Object.values(av.fixo).includes(Number(k))) delete av.tomSo[k]; });
    nums.forEach(n => {                            // à mão primeiro
      const f = av.fixo[n];
      if (f === 0) tonsNovos[n] = 0;
      else if (typeof f === 'number' && f > 0) { (tons[f] = tons[f] || []).push(n); tonsNovos[n] = f; }
    });
    /* O ROLO DE REFERÊNCIA, COMO ERA ANTES (12/11/2026, Junior: "opção A,
       comparado com rolo de referência, valor real"). A conferência a olho é
       feita encostando cada rolo no rolo de referência do tom; a regra de antes
       fazia a mesma coisa, e por isso o olho concordava. O "todas casam com
       todas" (e a média de todas as medições) não é o que a conferência faz.
       Então, como até 03/11, mas calculado das medições guardadas — sem rodada
       nem repescagem à mão:
         · Tom 1: a referência é a amostra de MENOR número (a 1ª medição dela);
           entra no tom toda amostra com ALGUMA medição aprovada contra ela
           ("uma aprovação basta");
         · Tom 2: a referência é a menor que sobrou, e assim por diante, até
           todas terem tom — o que antes era a repescagem;
         · aprovada = a regra das Configurações (dE E dL*, da*, db*), pela
           diferença REAL, sem arredondar.
       Duas amostras do mesmo tom casam com a referência; entre si podem ficar
       até o dobro da tolerância — é a regra que a conferência visual validava. */
    let resta = nums.filter(n => !(n in tonsNovos));
    const chavesUsadas = new Set(Object.keys(tons).map(Number));
    while (resta.length) {
      const ref = resta[0];
      const membros = resta.filter(n => n === ref || aprovadaContraRef(n, ref, av));
      let k = 1; while (chavesUsadas.has(k)) k++; chavesUsadas.add(k);
      tons[k] = membros.slice();
      membros.forEach(n => { tonsNovos[n] = k; });
      resta = resta.filter(n => !membros.includes(n));
    }
    // Amostra marcada no cadastro antigo e sem medição não some do mapa.
    Object.keys(av.tons || {}).forEach(n => { if (!(n in cores) && av.fixo[n] === 0) tonsNovos[n] = 0; });
    /* SEM SALTO NA NUMERAÇÃO (12/11/2026, Junior: "não deve haver incoerência
       numérica ordinal. Na medição o tom 2 foi pulado"). A chave interna do tom
       podia ficar com buraco (tom fixado à mão, "tom novo", tom que esvaziou), e
       cada tela tinha de lembrar de converter chave em posição. Agora as chaves
       SÃO as posições: renumeradas 1, 2, 3… pela menor amostra de cada tom, e os
       tons fixados à mão e os "só esta amostra" acompanham a renumeração. */
    const vivos = Object.keys(tons).map(Number).filter(k => tons[k].length)
      .sort((p, q) => Math.min(...tons[p]) - Math.min(...tons[q]));
    const novaChave = {};
    vivos.forEach((k, i) => { novaChave[k] = i + 1; });
    Object.keys(tonsNovos).forEach(n => { if (tonsNovos[n] > 0) tonsNovos[n] = novaChave[tonsNovos[n]]; });
    Object.keys(av.fixo).forEach(n => {
      const f = av.fixo[n];
      if (typeof f === 'number' && f > 0) { if (novaChave[f]) av.fixo[n] = novaChave[f]; else delete av.fixo[n]; }
    });
    if (av.tomSo) {
      const so = {};
      Object.keys(av.tomSo).forEach(k => { if (novaChave[k]) so[novaChave[k]] = true; });
      av.tomSo = so;
    }
    av.tons = tonsNovos;
    av.ref = {};
    vivos.forEach(kAntigo => {
      const k = novaChave[kAntigo];
      const n = Math.min(...tons[kAntigo]);
      const m0 = (av.medidas[n] || [])[0] || {};
      av.ref[k] = { medId: m0.id, amostra: n, nome: nomeAm(n), lab: m0.lab || cores[n], em: m0.em };
    });
  }
  // Uma avaliação gravada no modelo de rodadas vira o modelo novo sem perder medições:
  // o "Não usar" de antes fica fixo; o resto é reclassificado pelos dados.
  function avNormalizar(av) {
    if (!av || av.modelo === 'dados') return;
    av.fixo = av.fixo || {};
    Object.keys(av.tons || {}).forEach(n => { if (av.tons[n] === 0) av.fixo[n] = 0; });
    av.rodada = 1; av.fila = null; av.fimRodada = false; av.modelo = 'dados';
    if (av.n == null) { const f = faltamMedir(av); av.n = f.length ? f[0] : null; }
    classificar(av);
  }
  // A amostra nova contra cada tom: compatível com todas dele? e o maior dE.
  function encaixeNosTons(n, av = S.av) {
    if (!corDaAmostra(n, av)) return [];
    // TODOS os tons, o da própria amostra inclusive, cada um pelo ROLO DE
    // REFERÊNCIA dele: a amostra n passa contra a referência?
    return tonsDoRelatorio(av).map(t => {
      const proprio = av.tons[n] === t.k;
      if (t.refN === n) return { k: t.k, ok: true, pior: 0, proprio, ehRef: true, so: t.membros.length === 1, refN: n };
      const m = melhorContraRef(n, t.refN, av);
      return { k: t.k, ok: !!(m && m.aprovado), pior: m ? m.d : 0, proprio, ehRef: false, so: false, refN: t.refN };
    });
  }

  /* A RIBANA CONTRA O CORPO (12/11/2026, Junior: "quando concluir a fase corpo,
     na fase ribana as amostras devem ser comparadas com as amostras da fase
     corpo e determinar qual amostra de ribana é compatível com as amostras da
     fase corpo").
     A ribana é costurada no corpo: o que importa não é a ribana casar com outra
     ribana, é casar com o CORPO em que ela vai. Então, na fase de ribana que
     tem fase de cor (faseDeCorDaRibana: CM.LISA → Corpo, CM.REC → Corpo 1…), cada
     amostra de ribana — pela média — é comparada com CADA amostra medida do
     corpo, pela mesma tolerância das Configurações, só sim/não. Ela serve ao
     Tom k do corpo quando é aprovada contra o ROLO DE REFERÊNCIA desse tom
     (desde a opção A, 12/11/2026; antes: contra todas as amostras do tom). Pode servir a mais de um tom; se não serve a
     nenhum, é incompatível com o corpo ('x'). O tom gravado na ribana é o
     PRIMEIRO tom do corpo que ela serve — assim "Tom 1 da ribana" é a ribana do
     Tom 1 do corpo. Sem nenhuma amostra de corpo medida, a ribana volta a ser
     classificada sozinha, como antes. */
  function faseDaAv(av) {
    if (!S.avOS || !S.fases) return null;
    const i = Object.keys(S.avOS.fases).find(k => S.avOS.fases[k] === av);
    return i == null ? null : S.fases[Number(i)] || null;
  }
  function corpoDaRibana(av) {
    const f = faseDaAv(av);
    if (!f || !(f.corDe >= 0)) return null;
    const corpo = S.avOS.fases[f.corDe];
    if (!corpo || !numerosDasAmostras(corpo).some(n => corDaAmostra(n, corpo))) return null;
    return { f, fCorpo: S.fases[f.corDe], corpo };
  }
  // Ribana n contra o corpo: cada amostra do corpo e cada tom do corpo.
  function ribanaContraCorpo(n, av, ctx = corpoDaRibana(av)) {
    if (!ctx || !corDaAmostra(n, av)) return null;
    const corpo = ctx.corpo;
    // A ribana contra o ROLO DE REFERÊNCIA de cada tom do corpo (alguma medição
    // dela aprovada contra a 1ª medição da referência), como as amostras do corpo.
    const tons = tonsDoRelatorio(corpo).map((t, i) => {
      const m = melhorContra(av.medidas[n], labRef(t.refN, corpo));
      return { k: t.k, ordem: i + 1, membros: t.membros, refN: t.refN, ok: !!(m && m.aprovado), r: m ? m.r : null };
    });
    const porAmostra = tons.filter(t => t.r).map(t => ({ x: t.refN, ok: t.ok, r: t.r, tom: t.ordem }));
    return { porAmostra, tons, serve: tons.filter(t => t.ok).map(t => t.ordem) };
  }
  function classificarRibana(av, ctx) {
    av.fixo = av.fixo || {};
    const tons = {};
    av.serve = {};
    numerosDasAmostras(av).filter(n => corDaAmostra(n, av)).forEach(n => {
      const f = av.fixo[n];
      const rc = ribanaContraCorpo(n, av, ctx);
      av.serve[n] = rc ? rc.serve : [];
      if (f === 0) { tons[n] = 0; return; }
      if (typeof f === 'number' && f > 0) { tons[n] = f; av.serve[n] = [f]; return; }
      tons[n] = av.serve[n].length ? av.serve[n][0] : 'x';
    });
    av.tons = tons;
    av.ref = {};
    av.contraCorpo = true;
  }
  /* TODAS AS FASES EM DIA (12/11/2026, caso OS 0715). Só a ribana era refeita
     aqui; o corpo só se reclassificava com uma medição nova nele ou ao entrar
     na fase. Com a tela na Gola, o corpo seguiu com os tons gravados pela regra
     anterior — e mudar a regra, ou as tolerâncias nas Configurações, deixava os
     tons velhos na tela, no recado e no exportar. Agora toda fase é refeita a
     cada desenho (o corpo antes da ribana, que depende dele), e o resultado é
     guardado quando muda. A conta é pequena: dezenas de amostras. */
  function sincronizarRibanas() {
    if (!S.avOS || !S.fases) return;
    const antes = JSON.stringify(Object.keys(S.avOS.fases).map(i => S.avOS.fases[i] && S.avOS.fases[i].tons));
    S.fases.filter(f => !(f.corDe >= 0)).forEach(f => { if (S.avOS.fases[f.i]) { avNormalizar(S.avOS.fases[f.i]); classificar(S.avOS.fases[f.i]); } });
    S.fases.filter(f => f.corDe >= 0).forEach(f => { if (S.avOS.fases[f.i]) { avNormalizar(S.avOS.fases[f.i]); classificar(S.avOS.fases[f.i]); } });
    const depois = JSON.stringify(Object.keys(S.avOS.fases).map(i => S.avOS.fases[i] && S.avOS.fases[i].tons));
    if (antes !== depois) avGuardar();
  }
  const ehIncompativelCorpo = (n, av = S.av) => tomDe(n, av) === 'x';

  // A referência do tom como "leitura", para a comparação da tela (barra, blocos, lista).
  function refComoLeitura(k) {
    const r = S.av && S.av.ref[k];
    if (!r) return null;
    const ja = S.leituras.find(x => x.id === r.medId);
    return ja || { id: r.medId, nome: r.nome, tipo: 'padrao', lab1: r.lab, lab2: r.lab2 || null, espectro: [], horario: new Date(r.em) };
  }

  /* A OS EM AVALIAÇÃO SOBREVIVE AO F5 (20/10/2026). Ela só se escolhia pelo
     "⋯ → avaliar amostras" e ficava só na memória: depois de um F5 a tela
     voltava no modo sem OS (Medir padrão / Medir amostra) e o quadro da
     avaliação, com o número da amostra, sumia. Agora a OS fica guardada neste
     navegador até o "Medir sem OS". */
  const CHAVE_OS = 'gos:colorimetro:os';
  function abrirOS(o) {
    const ehOP = !!listaOP().find(x => x.id === o.id);
    const numero = ehOP ? opNum(o) : (o.os || '');
    S.os = { id: o.id, numero, doc: ehOP ? 'OP' : 'OS', nome: (ehOP ? 'OP ' : 'OS ') + numero, rotulo: ehOP ? rotuloOP(o) : rotuloOS(o) };
    S.avOS = avCarregar(o.id);
    S.fases = fasesDaOS(o.id);
    // Na OP o total de amostras de cada fase já vem do lote: uma por bobina.
    // Só preenche o que ainda não foi escolhido (quem mudou o total, manda).
    if (ehOP) S.fases.forEach(f => {
      if (!(f.total > 0)) return;
      const av = S.avOS.fases[f.i] = S.avOS.fases[f.i] || avNovo(o.id);
      if (!(Number(av.total) > 0)) av.total = f.total;
    });
    let i = Math.min(S.avOS.faseAtual || 0, S.fases.length - 1);
    // Parou numa fase que não precisa de avaliação: vai para a primeira que falta.
    if (S.fases[i] && S.fases[i].semAval) {
      const j = S.fases.findIndex(f => !faseFeita(S.avOS, f.i));
      i = j >= 0 ? j : Math.max(0, S.fases.findIndex(f => !f.semAval));
    }
    avUsarFase(i);
    try { localStorage.setItem(CHAVE_OS, o.id); } catch (e) { /* sem storage */ }
  }
  function restaurarOS() {
    if (S.os || S.osRestaurada) return;
    if (!listaOS().length && !listaOP().length) return; // os dados ainda não chegaram: tenta no próximo desenho
    S.osRestaurada = true;
    let id = null;
    try { id = localStorage.getItem(CHAVE_OS); } catch (e) { /* sem storage */ }
    const d = id && acharDoc(id);
    if (!d) return;
    abrirOS(d.o);
    S.status = S.dev ? avDica() : `${S.os.nome} escolhida — conecte o aparelho para medir`;
  }
  // Sem OS: a OS também se escolhe aqui, não só pelo "⋯ → avaliar amostras" da lista.
  function htmlEscolherOS() {
    const ordens = (typeof STATE !== 'undefined' && STATE && STATE.ordens || [])
      .filter(o => o && o.id && o.os)
      .slice().sort((a, b) => String(b.os).localeCompare(String(a.os), 'pt-BR', { numeric: true }))
      .slice(0, 80);
    const ops = listaOP().filter(p => p && p.id && p.status !== 'cancelada')
      .slice().sort((a, b) => (Number(b.numero) || 0) - (Number(a.numero) || 0)).slice(0, 80);
    if (!ordens.length && !ops.length) return '';
    return `<label class="cm-campo" style="max-width:520px;margin-bottom:10px"><span>Avaliar as amostras da OS ou da OP</span>
      <select onchange="if(this.value)Colorimetro.avaliarOS(this.value)">
        <option value="">— escolha a OS ou a OP —</option>
        ${ops.length ? `<optgroup label="Ordens de preparo de matéria-prima (OP)">${ops.map(p => `<option value="${escH(p.id)}">${escH(rotuloOP(p))}</option>`).join('')}</optgroup>` : ''}
        ${ordens.length ? `<optgroup label="Ordens de serviço (OS)">${ordens.map(o => `<option value="${escH(o.id)}">${escH(rotuloOS(o))}</option>`).join('')}</optgroup>` : ''}</select></label>`;
  }
  function avaliarOS(id) {
    const d = acharDoc(id);
    if (!d) return;
    abrirOS(d.o);
    S.janela = null;
    S.aviso = '';
    if (typeof goto === 'function') goto('colorimetro');
    status(S.dev ? avDica() : `${S.os.nome} escolhida — conecte o aparelho para medir`);
  }
  function encerrarOS() {
    try { localStorage.removeItem(CHAVE_OS); } catch (e) { /* sem storage */ }
    S.os = null; S.av = null; S.avOS = null; S.fases = null; S.padrao = null; S.atual = null;
    if (S.janela === 'rel') S.janela = null;
    status('Medições sem OS');
  }

  function avDica() {
    if (!S.av) return '';
    const fase = S.fases && S.fases[S.avOS.faseAtual];
    const pre = fase && S.fases.length > 1 ? fase.rotulo + ' — ' : '';
    if (S.avOS.concluidas[S.avOS.faseAtual]) return pre + 'fase concluída. Para repetir, escolha-a em "Reavaliar a fase".';
    return pre + avDicaFase();
  }
  function avDicaFase() {
    const av = S.av;
    if (av.n == null) {
      const t = totalDe(av);
      return t ? `As ${t} amostras desta fase foram medidas. Confira os tons e conclua a fase (ou escolha o número de uma amostra para medir de novo).`
        : 'Escolha o número da amostra para medir.';
    }
    return `Meça a Amostra ${av.n}: a cor fica guardada e o programa a compara com todas as já medidas desta fase. Depois clique em Trocar.`;
  }
  // Não há mais "a reavaliar" (sem repescagem); fica para avaliações gravadas antes.
  const avPendentes = (av = S.av) => numerosDasAmostras(av).filter(n => tomDe(n, av) === 'r');

  // Uma medição chegou com a tela ligada a uma OS (chamado por chegou()).
  // F5, F6 e o botão do aparelho fazem a mesma coisa aqui: medem a amostra atual.
  function avMedicao(l) {
    const av = S.av;
    l.osId = S.os.id; l.osNumero = S.os.numero; l.rodada = 1; l.tipo = 'amostra';
    l.fase = S.avOS.faseAtual; l.faseNome = (S.fases[S.avOS.faseAtual] || {}).rotulo || '';
    if (av.n == null) return avDicaFase();
    const n = av.n;
    l.amostraN = n;
    l.nome = `${cfg.nome_amostra} ${n}`;
    const em = l.horario.toISOString();
    const lista = av.medidas[n] = av.medidas[n] || [];
    const antes = tomDe(n);
    lista.push({ id: l.id, lab: l.lab1, rodada: 1, de: 0, aprovado: true, em });
    classificar(av);
    const k = tomDe(n);
    const ref = estaNumTom(n) ? refComoLeitura(k) : null;
    if (ref && ref.id !== l.id) {
      const r = comparar(ref, l);
      lista[lista.length - 1].de = r.des[cfg.formula];
      lista[lista.length - 1].aprovado = r.aprovado;
      l.refId = ref.id;
    }
    S.padrao = ref;
    avGuardar();
    if (av.contraCorpo) {
      const rc = ribanaContraCorpo(n, av);
      const ctx = corpoDaRibana(av);
      const nomeCorpo = ctx ? ctx.fCorpo.nome : 'corpo';
      const prox2 = proximoNumero(n);
      const dep = prox2 != null ? ` Clique em Trocar para passar à Amostra ${prox2}.` : '';
      if (k === 0) return `Amostra ${n} medida — está marcada "Não usar" no relatório.${dep}`;
      if (rc && rc.serve.length) return `✔ Amostra ${n} aprovada contra a referência do ${rc.serve.map(o => 'Tom ' + o).join(' e do ')} do ${nomeCorpo} → serve ao ${rc.serve.map(o => 'Tom ' + o).join(' e ao ')}.${dep}`;
      return `✘ Amostra ${n} reprovada contra a referência de todos os tons do ${nomeCorpo}: incompatível.${dep}`;
    }
    const tonsAgora = tonsDoRelatorio();
    const tAqui = tonsAgora.find(t => t.k === k);
    const ordem = tonsAgora.indexOf(tAqui) + 1;
    const prox = proximoNumero(n);
    const depois = prox != null ? ` Clique em Trocar para passar à Amostra ${prox}.` : '';
    if (k === 0) return `Amostra ${n} medida — está marcada "Não usar" no relatório.${depois}`;
    const mudou = typeof antes === 'number' && antes > 0 && antes !== k ? ' (mudou de tom com esta medição)' : '';
    // O rolo de referência: a amostra aprovada contra a referência entra no tom
    // dela; a que não passa em nenhuma referência anterior vira a referência do
    // tom seguinte (o que antes era a repescagem).
    if (tAqui && tAqui.refN === n) {
      const anteriores = tonsAgora.filter(t => t.refN < n).map(t => t.refN);
      return anteriores.length
        ? `✘ Amostra ${n} reprovada contra a referência ${anteriores.length > 1 ? 'dos tons anteriores (amostras ' + anteriores.join(', ') + ')' : 'do Tom 1 (Amostra ' + anteriores[0] + ')'} → é a referência do Tom ${ordem}${mudou}.${depois}`
        : `Amostra ${n} é a referência do Tom ${ordem}${mudou}.${depois}`;
    }
    return `✔ Amostra ${n} aprovada contra a referência (Amostra ${tAqui ? tAqui.refN : '?'}) → Tom ${ordem}${mudou}.${depois}`;
  }


  // O número seguinte: na 1ª rodada, o próximo inteiro (amostras novas); nas
  // seguintes, a próxima da fila de reavaliação ainda não medida nesta rodada —
  // dando a volta, porque a referência escolhida pode estar no meio da fila
  // (Amostra 5 com 3 e 4 antes dela).
  /* O TOTAL DE AMOSTRAS (12/11/2026, Junior: "no campo do colorímetro insira
     campo onde usuário determina quantas amostras serão avaliadas no total").
     Sem total, a 1ª rodada segue infinita (n + 1), como sempre foi. Com total,
     ela anda só de 1 até ele — pulando as já medidas — e termina sozinha na
     última, e a tela diz "Nº 3 de 10" e quais ainda faltam. O total é da FASE
     (cada tecido tem os seus rolos); o último digitado vira o padrão da OS para
     as fases que ainda não têm o seu. */
  // av.total = 0 é "sem limite" escolhido na fase; sem o campo, vale o padrão da OS.
  const totalDe = (av = S.av) => Number(av && av.total != null ? av.total : (S.avOS && S.avOS.total)) || 0;
  const faltamMedir = (av = S.av) => {
    const t = totalDe(av);
    return t ? Array.from({ length: t }, (_, i) => i + 1).filter(x => !(av.medidas[x] || []).length) : [];
  };
  function avDefinirTotal(v) {
    const t = parseInt(v, 10);
    if (!S.av || !S.avOS) return;
    if (!(t > 0)) { S.av.total = 0; avGuardar(); status('Total de amostras apagado: a rodada segue sem fim, amostra a amostra.'); return; }
    S.av.total = t; S.avOS.total = t;
    // Total menor que a amostra atual da 1ª rodada: a rodada acaba aqui.
    if (S.av.rodada === 1 && !S.av.fila && S.av.n != null && S.av.n > t) { S.av.n = null; S.av.fimRodada = true; }
    else if (S.av.rodada === 1 && !S.av.fila && S.av.n == null && S.av.fimRodada && faltamMedir().length) { S.av.n = faltamMedir()[0]; S.av.fimRodada = false; }
    avGuardar();
    const f = faltamMedir();
    status(`Total desta fase: ${t} ${t === 1 ? 'amostra' : 'amostras'}.` + (f.length ? ` Faltam medir: ${f.join(', ')}.` : ' Todas já foram medidas.'));
  }

  function proximoNumero(n) {
    const av = S.av;
    if (av.rodada === 1 && !av.fila) {
      const t = totalDe(av);
      if (!t) return n + 1;
      const falta = faltamMedir(av).filter(x => x !== n);
      return falta.find(x => x > n) || falta[0] || null;
    }
    const fila = av.fila || [];
    const medida = x => (av.medidas[x] || []).some(m => m.rodada === av.rodada);
    const i = fila.indexOf(n);
    const depois = i >= 0 ? fila.slice(i + 1) : fila.filter(x => x > n);
    return depois.find(x => !medida(x)) || fila.find(x => x !== n && !medida(x)) || null;
  }

  function avTrocar() {
    const av = S.av;
    if (!av || av.n == null) return;
    const n = av.n;
    // Sem repescagem: trocar só passa à amostra seguinte; o tom já foi decidido na medição.
    let msg = (av.medidas[n] || []).length ? '' : `Amostra ${n} ficou sem medição. `;
    const prox = proximoNumero(n);
    if (prox == null) {
      av.n = null; av.fimRodada = true;
      msg += avDica();
    } else {
      av.n = prox;
      msg += `Agora é a Amostra ${prox}.`;
    }
    S.atual = null;
    avGuardar();
    status(msg);
  }
  function avIrPara(v) {
    const n = parseInt(v, 10);
    if (!S.av || !(n > 0)) return;
    const t = totalDe();
    if (t && n > t) { S.aviso = `Esta fase tem ${t} ${t === 1 ? 'amostra' : 'amostras'} no total — aumente o total para ir à Amostra ${n}.`; desenhar(); return; }
    S.av.n = n; S.av.fimRodada = false;
    S.atual = null;
    S.padrao = estaNumTom(n) ? refComoLeitura(tomDe(n)) : null;
    avGuardar();
    const ordem = tonsDoRelatorio().findIndex(t => t.k === tomDe(n)) + 1;
    status(`Agora é a Amostra ${n}${ordem ? ` (já medida: Tom ${ordem}; medir de novo soma à média dela)` : tomDe(n) === 0 ? ' (marcada "Não usar")' : ''}.`);
  }
  // Concluir a fase atual e passar para a seguinte (as fases andam em ordem).
  function avConcluirFase() {
    const c = S.avOS;
    if (!c) return;
    const i = c.faseAtual;
    if (!numerosDasAmostras().length) { S.aviso = 'Meça ao menos uma amostra desta fase antes de concluí-la.'; desenhar(); return; }
    const falta = faltamMedir();
    if (falta.length && !S._concluirMesmoAssim) {
      perguntar({ titulo: `Concluir a ${S.fases[i].rotulo} com amostras sem medir?`,
        texto: `O total desta fase é ${totalDe()} e ${falta.length === 1 ? 'a amostra' : 'as amostras'} ${falta.join(', ')} ainda não ${falta.length === 1 ? 'foi medida' : 'foram medidas'}.`,
        sim: 'Concluir mesmo assim', executar: () => { S._concluirMesmoAssim = true; try { avConcluirFase(); } finally { S._concluirMesmoAssim = false; } } });
      return;
    }
    c.concluidas[i] = true;
    const prox = S.fases.findIndex((f, j) => j > i && !faseFeita(c, j));
    if (prox >= 0) avUsarFase(prox);
    avGuardar();
    status(prox >= 0 ? `${S.fases[i].rotulo} concluída. Agora é a ${S.fases[prox].rotulo}: comece pela Amostra 1.`
      : `${S.fases[i].rotulo} concluída. Todas as fases foram avaliadas — abra o Relatório de tonalidades.`);
  }
  // Repetir uma fase: a pessoa aponta QUAL. A avaliação dela recomeça do zero.
  function avReavaliarFase(v) {
    const i = parseInt(v, 10);
    const c = S.avOS;
    if (!c || !(i >= 0) || !S.fases[i]) return;
    perguntar({ titulo: `Reavaliar a ${S.fases[i].rotulo}`,
      texto: 'Os tons desta fase serão refeitos desde a Amostra 1. As medições antigas continuam no histórico deste computador, e o relatório vai dizer que a fase foi reavaliada.',
      sim: 'Reavaliar esta fase', executar: () => {
        c.fases[i] = avNovo(c.osId);
        delete c.concluidas[i];
        c.reav[i] = (c.reav[i] || 0) + 1;
        avUsarFase(i);
        avGuardar();
        status(`Reavaliação da ${S.fases[i].rotulo}: meça a Amostra 1; os tons se formam pelas cores medidas.`);
      } });
  }

  /* LIMPAR A AVALIAÇÃO (03/11/2026, Jones: "insira função de limpar a avaliação
     de cores do colorímetro, para que o usuário possa recomeçar a avaliação
     completa, sem nenhuma amostra avaliada como referência"). O "Reavaliar a
     fase" refaz uma fase concluída; este apaga a avaliação da OS inteira —
     todas as fases, tons, referências, concluídas e reavaliações — e volta à
     Amostra 1 da primeira fase que precisa de avaliação, sem referência. A
     lista da tela também é limpa, para nenhuma leitura antiga servir de padrão.
     As medições continuam no histórico deste computador (IndexedDB), e o que
     já foi gravado nas observações da OS pelo relatório não é tocado. */
  function avLimpar() {
    const c = S.avOS;
    if (!c || !S.os) return;
    perguntar({ titulo: `Limpar a avaliação da ${S.os.nome || "OS " + S.os.numero}`,
      texto: 'Todas as fases voltam ao começo: tons, referências, amostras medidas, fases concluídas e reavaliações são apagados, e a avaliação recomeça na Amostra 1, sem nenhuma referência. A lista da tela também é limpa. As medições continuam no histórico deste computador, e o que já foi gravado nas observações da OS não muda.',
      sim: 'Limpar a avaliação', executar: () => {
        S.avOS = avNovoOS(c.osId);
        const j = S.fases.findIndex(f => !f.semAval);
        avUsarFase(j >= 0 ? j : 0);
        S.leituras = []; S.cont = { padrao: 0, amostra: 0 };
        S.padrao = null; S.atual = null;
        if (S.janela === 'rel') S.janela = null;
        avGuardar();
        status(`Avaliação limpa. ${avDica()}`);
      } });
  }

  /* ---------------- o relatório ---------------- */
  // Os tons em uso, na ordem dos números das rodadas.
  function tonsDoRelatorio(av = S.av) {
    const nums = numerosDasAmostras(av);
    // Os tons na ordem da menor amostra de cada um: o Tom 1 é sempre o da amostra mais baixa.
    const menorDe = k => Math.min(...nums.filter(n => tomDe(n, av) === k));
    const ks = [...new Set(nums.map(n => tomDe(n, av)).filter(t => typeof t === 'number' && t > 0))].sort((a, b) => menorDe(a) - menorDe(b));
    return ks.map(k => {
      const membros = nums.filter(n => tomDe(n, av) === k);
      const ref = av.ref[k];
      const refN = ref && ref.amostra != null ? ref.amostra : Math.min(...membros);
      // Cada amostra contra o ROLO DE REFERÊNCIA do tom, pela melhor medição dela:
      // pior = o maior dE contra a referência; ruins = as que não passam contra
      // ela (só acontece com tom marcado à mão no relatório).
      const melhor = n => { const c = corDaAmostra(n, av); return c ? { lab: c } : null; };
      const ruins = [];
      let pior = 0;
      membros.forEach(n => {
        if (n === refN) return;
        const m = melhorContraRef(n, refN, av);
        if (!m) return;
        pior = Math.max(pior, m.d);
        if (!m.aprovado) ruins.push([n, m.d]);
      });
      return { k, membros, ref, refN, pior, ruins, melhor };
    });
  }
  const nomeAm = n => `${cfg.nome_amostra} ${n}`;
  // As fases que já têm medição, na ordem da OS.
  const fasesComMedicao = () => (S.fases || []).filter(f => !f.semAval && S.avOS.fases[f.i] && numerosDasAmostras(S.avOS.fases[f.i]).length);
  function marcasDaFase(f) {
    const r = S.avOS.reav[f.i];
    return [r ? `reavaliada${r > 1 ? ' ' + r + 'x' : ''}` : '', S.avOS.concluidas[f.i] ? '' : 'em andamento'].filter(Boolean);
  }

  /* O RECADO NAS OBSERVAÇÕES, SIMPLIFICADO (20/10/2026, Junior: "deve aparecer
     de forma simplificada o conjunto de números das amostras compatíveis, por
     exemplo, como título do relatório: Tecidos compatíveis em cor, Tecido: Fase
     corpo (1,2,3,4,5,6 - ribana 1)").
     Quem lê a folha quer saber QUAIS amostras podem ir juntas, e não o dE, a
     referência nem a fórmula — isso fica na tela do relatório. Então: o título,
     e por fase do enfesto (na ordem dela) o tecido e os números de cada grupo
     compatível entre parênteses; mais de um grupo, um parêntese por grupo. As
     que ainda esperam reavaliação ou foram marcadas "Não usar" vão no fim da
     fase, curtas, porque ficar de fora também é informação para o corte.
       Tecidos compatíveis em cor
       Corpo · Malha Algodão Vermelho - Tom 1 (1,2,3,4,5,6) - Gola · Ribana - Tom 1 (1)
     COM O TOM ESCRITO (26/10/2026, Junior: "Corpo Parte 1 · Malha Algodão Preto -
     Tom 1 (1,2,3,5,6,7,8,9,10 - ribana 2,3) Tom 2 (4 - ribana 2,3), ribana 1 não
     usar"): cada grupo diz de que tom é; a ribana de um tom só vai com TODOS os
     tons do corpo (ela casa com a fase inteira); e o número vem antes do
     "não usar" / "a reavaliar". */
  // A ribana conjugada nesta fase (já medida), ou nenhuma.
  const ribanasDe = (f, medidas) => medidas.filter(r => r.corDe === f.i);
  function textoFaseCurto(f, medidas = fasesComMedicao()) {
    const av = S.avOS.fases[f.i];
    const rib = ribanasDe(f, medidas).map(r => ({ r, av: S.avOS.fases[r.i], tons: tonsDoRelatorio(S.avOS.fases[r.i]) }));
    const tons = tonsDoRelatorio(av);
    // Ribana comparada com o corpo (12/11/2026): ela entra em CADA tom do corpo
    // a que serve, e não passa do número de tons do corpo.
    const nGrupos = Math.max(tons.length, ...rib.map(x => x.av.contraCorpo ? 0 : x.tons.length));
    // Tom a tom: Tom k (corpo - ribana); a ribana com mais tons que o corpo sai
    // só ela, e a de um tom só entra em todos (ribana sem corpo medido).
    const tomDaRibana = (x, j) => {
      if (x.av.contraCorpo) {
        const m = numerosDasAmostras(x.av).filter(n => tomDe(n, x.av) !== 0 && (x.av.serve[n] || []).includes(j + 1));
        return m.length ? { membros: m } : null;
      }
      return x.tons.length === 1 ? x.tons[0] : x.tons[j];
    };
    const grupos = Array.from({ length: nGrupos }, (_, j) => {
      // SEM O "ref N:" (12/11/2026, Junior: "retire do relatório auto produzido o
      // texto ref1, ref2, ref3, pois é redundante, uma vez que no texto já tem
      // escrito tom 1, tom 2, tom 3"). A referência de cada tom continua na tela
      // do relatório do Colorímetro; o texto da folha leva só as amostras.
      const partes = [tons[j] ? tons[j].membros.join(',') : '',
        ...rib.map(x => tomDaRibana(x, j) ? `ribana ${tomDaRibana(x, j).membros.join(',')}` : '')].filter(Boolean);
      // O número que se escreve é a posição do tom (Tom 1, 2, 3…), não a chave interna.
      const k = j + 1;
      return `Tom ${k} (${partes.join(' - ')})`;
    }).join(' ');
    const sobra = (a, rotulo) => {
      // Incompatível com o corpo ('x') também não vai: sai junto do "não usar".
      const pend = avPendentes(a), fora = numerosDasAmostras(a).filter(n => tomDe(n, a) === 0 || tomDe(n, a) === 'x');
      return (pend.length ? `, ${rotulo}${pend.join(',')} a reavaliar` : '') + (fora.length ? `, ${rotulo}${fora.join(',')} não usar` : '');
    };
    const nome = [f.nome, f.tecido].filter(Boolean).join(' · ');
    const reav = S.avOS.reav[f.i] || rib.some(x => S.avOS.reav[x.r.i]);
    return `${nome}${reav ? ' [reavaliada]' : ''} - ${grupos || '(nenhuma)'}`
      + sobra(av, '') + rib.map(x => sobra(x.av, 'ribana ')).join('');
  }
  // Um recado só, com todas as fases medidas, na ordem do enfesto; a ribana
  // conjugada vai dentro da fase de cor dela, não sozinha.
  function textoRelatorio() {
    const medidas = fasesComMedicao();
    const soltas = medidas.filter(f => !(f.corDe >= 0 && medidas.some(c => c.i === f.corDe)));
    return 'Tecidos compatíveis em cor\n' + soltas.map(f => textoFaseCurto(f, medidas)).join(' - ');
  }
  /* "TOM NOVO" (20/10/2026, Junior: "corrija as opções na coluna Tom do
     relatório de tonalidade, para que tenha como escolher apenas uma amostra
     como sendo Tom 2"). O seletor só oferecia os tons que já tinham rodada:
     com todas as amostras aprovadas no Tom 1, não havia Tom 2 para escolher, e
     separar UMA amostra exigia medir de novo. Agora há sempre o tom seguinte
     ao último, e a amostra escolhida para ele é a referência dele (pela última
     medição dela) — o tom nasce com ela sozinha. Reavaliar depois abre a
     rodada seguinte a esse, sem reaproveitar o número. */
  function tomNovoDe(av) {
    const usados = Object.values(av.tons).filter(t => typeof t === 'number' && t > 0);
    return Math.max(av.rodada || 1, ...Object.keys(av.ref).map(Number), ...usados) + 1;
  }
  function relMudar(i, n, v) {
    const av = S.avOS && S.avOS.fases[i];
    if (!av) return;
    // O que se escolhe aqui fica FIXO (av.fixo) e vale por cima da classificação
    // pelos dados; "Automático" devolve a amostra à classificação.
    av.fixo = av.fixo || {};
    if (v === 'auto' || v === '') delete av.fixo[n];
    else if (v === 'novo' && !av.contraCorpo) { const k = tomNovoDe(av); av.fixo[n] = k; (av.tomSo = av.tomSo || {})[k] = true; }
    else av.fixo[n] = parseInt(v, 10) || 0;
    classificar(av);
    avGuardar();
    desenhar();
  }
  async function relSalvar() {
    if (!S.avOS || !S.os) return;
    sincronizarRibanas();
    if (!fasesComMedicao().some(f => tonsDoRelatorio(S.avOS.fases[f.i]).length)) { S.aviso = 'Nenhuma amostra está num tom ainda.'; desenhar(); return; }
    if (S.os.doc !== 'OP' && typeof salvarObsNota !== 'function') { S.aviso = 'Não achei a gravação das observações da OS.'; desenhar(); return; }
    // O EDITAR DO COLORÍMETRO (21/10/2026): gravar o relatório de cor na OS é
    // de quem o admin marcar em Configurações → contas de acesso.
    if (typeof currentRole !== 'undefined' && currentRole !== 'admin' && typeof contaTemAcesso === 'function'
        && !contaTemAcesso('colorimetro', typeof _obsQuemSou === 'function' ? _obsQuemSou() : '')) {
      S.aviso = 'Seu acesso não inclui gravar o relatório do Colorímetro na OS — peça ao admin.'; desenhar(); return;
    }
    // A OP guarda o relatório na folha dela (ver salvarRelatorioCorOP em app.js).
    if (S.os.doc === 'OP') {
      if (typeof salvarRelatorioCorOP !== 'function') { S.aviso = 'Não achei a gravação do relatório da OP.'; desenhar(); return; }
      let ok = false;
      try { ok = await salvarRelatorioCorOP(S.os.id, textoRelatorio()); } catch (e) { S.aviso = 'Não deu para salvar: ' + e.message; desenhar(); return; }
      if (!ok) { S.aviso = `O programa não gravou o relatório na ${S.os.nome}.`; desenhar(); return; }
      if (typeof window._diarioAnotar === 'function') window._diarioAnotar('colorimetro', `relatório de cor da ${S.os.nome}`, 'colorimetro');
      S.relSalvo = { osId: S.os.id, em: new Date() };
      S.aviso = '';
      status(`Relatório salvo na folha da ${S.os.nome}`);
      return;
    }
    const o = (STATE.ordens || []).find(x => x.id === S.os.id);
    if (!o) { S.aviso = `A OS ${S.os.numero} não existe mais.`; desenhar(); return; }
    const notas = () => (Array.isArray(o.obsNotas) ? o.obsNotas.length : 0);
    const antes = notas();
    try { await salvarObsNota(o.id, '', textoRelatorio()); } catch (e) { S.aviso = 'Não deu para salvar: ' + e.message; desenhar(); return; }
    if (notas() <= antes) { S.aviso = 'O programa não gravou a observação (veja o aviso no canto da tela).'; desenhar(); return; }
    // O Avisar da área Colorímetro: a gravação vai para o diário de alterações.
    if (typeof window._diarioAnotar === 'function') window._diarioAnotar('colorimetro', `relatório de cor da OS ${o.os}`, 'colorimetro');
    S.relSalvo = { osId: o.id, em: new Date() };
    S.aviso = '';
    status(`Relatório salvo nas observações da OS ${o.os}`);
  }

  // As fases em fila, como chips: a atual destacada, as concluídas com ✔.
  function htmlFases() {
    if (!S.fases || S.fases.length < 2) return '';
    const c = S.avOS;
    const chips = S.fases.map(f => {
      if (f.semAval) return `<span class="cm-fase cm-fase-sem" title="As peças desta fase não são unidas a outras da mesma cor">${escH(f.rotulo)} <span class="cm-fraco">· não precisa avaliar</span></span>`;
      const atual = f.i === c.faseAtual, feita = !!c.concluidas[f.i];
      const av = c.fases[f.i];
      const tons = av ? tonsDoRelatorio(av).length : 0;
      return `<span class="cm-fase ${atual ? 'cm-fase-atual' : ''} ${feita ? 'cm-fase-feita' : ''}">${feita ? '✔ ' : ''}${escH(f.rotulo)}`
        + `${av && numerosDasAmostras(av).length ? ` <span class="cm-fraco">· ${tons} ${tons === 1 ? 'tom' : 'tons'}</span>` : ''}${c.reav[f.i] ? ' <span class="cm-fraco">· reavaliada</span>' : ''}</span>`;
    }).join('');
    const feitas = S.fases.filter(f => c.concluidas[f.i] && !f.semAval);
    return `<div class="cm-fases">${chips}</div>
      ${feitas.length ? `<div class="cm-botoes" style="margin-top:6px">
        <label class="cm-campo" style="max-width:340px"><span>Reavaliar a fase</span>
          <select onchange="if(this.value!=='')Colorimetro.avReavaliarFase(this.value); this.value=''">
            <option value="">— escolha a fase que vai ser reavaliada —</option>
            ${feitas.map(f => `<option value="${f.i}">${escH(f.rotulo)}</option>`).join('')}</select></label></div>` : ''}`;
  }

  // Ribana: para cada tom do corpo, as ribanas que servem a ele.
  function htmlRibanaPorTom(av, rib, fase) {
    const tonsCorpo = tonsDoRelatorio(rib.corpo);
    const nums = numerosDasAmostras(av).filter(n => corDaAmostra(n, av));
    const linhas = tonsCorpo.map((t, i) => {
      const servem = nums.filter(n => tomDe(n, av) !== 0 && (av.serve[n] || []).includes(i + 1));
      return `<b>${escH(rib.fCorpo.nome)} Tom ${i + 1}</b> (ref ${t.refN}: ${t.membros.join(', ')}): ${servem.length ? 'ribana ' + servem.join(', ') : '<span class="cm-fora">nenhuma ribana ainda</span>'}`;
    });
    const incomp = nums.filter(n => ehIncompativelCorpo(n, av)), nao = nums.filter(n => tomDe(n, av) === 0);
    const aviso = S.avOS.concluidas[rib.f.corDe] ? '' : `<br><span class="cm-fora">${escH(rib.fCorpo.nome)} ainda não concluído: a comparação usa as amostras medidas até agora.</span>`;
    return `<div class="cm-av-bloco"><div class="cm-av-rot">${S.fases.length > 1 ? escH(fase.rotulo) + ' · ' : ''}Contra os tons do ${escH(rib.fCorpo.nome)}</div>
      <div class="cm-peq">${linhas.join('<br>') || 'o corpo ainda não tem tom'}
      ${incomp.length ? `<br><span class="cm-fora">Incompatíveis com o ${escH(rib.fCorpo.nome)}: ${incomp.join(', ')}</span>` : ''}
      ${nao.length ? `<br><span class="cm-fora">Não usar: ${nao.join(', ')}</span>` : ''}${aviso}</div></div>`;
  }

  function htmlAvaliacao(aparelho = '', comparacao = '') {
    if (!S.os || !S.av) return '';
    const av = S.av;
    const con = !!S.dev, livre = !S.ocupado && !S.conectando;
    const concluida = !!S.avOS.concluidas[S.avOS.faseAtual];
    const fase = S.fases[S.avOS.faseAtual];
    const ultimaFase = !S.fases.some((f, j) => j > S.avOS.faseAtual && !faseFeita(S.avOS, j));
    const ms = av.n != null ? (av.medidas[av.n] || []) : [];
    const tons = tonsDoRelatorio();
    const ordemDe = k => tons.findIndex(t => t.k === k) + 1;
    // Ribana contra o corpo: os tons que contam são os do CORPO.
    const rib = av.contraCorpo ? corpoDaRibana(av) : null;
    const nomeCorpo = rib ? rib.fCorpo.nome : '';
    const rcAtual = rib && av.n != null && ms.length ? ribanaContraCorpo(av.n, av, rib) : null;
    const situacao = av.n == null || !ms.length ? '<span class="cm-fraco">ainda não medida</span>'
      : tomDe(av.n) === 0 ? '<span class="cm-fora">não usar</span>'
      : rib ? (rcAtual && rcAtual.serve.length ? `<span class="cm-ok">serve ao ${rcAtual.serve.map(o => 'Tom ' + o).join(', ')} do ${escH(nomeCorpo)}</span>`
        : `<span class="cm-fora">incompatível com o ${escH(nomeCorpo)}</span>`)
      : estaNumTom(av.n) ? `<span class="cm-ok">Tom ${ordemDe(tomDe(av.n))}</span>` : '';
    // A amostra atual contra cada tom formado: casa com todas dele? e o maior dE.
    const encaixe = rcAtual ? rcAtual.tons.map(t =>
      `<span class="${t.ok ? 'cm-ok' : 'cm-fora'}">${escH(nomeCorpo)} Tom ${t.ordem} ${t.ok ? '✔' : '✘'}</span>`).join(' · ')
      : av.n != null && ms.length && !rib ? encaixeNosTons(av.n).map(e =>
      e.ehRef ? `<span class="cm-ok">Tom ${ordemDe(e.k)} · é a referência</span>`
        : `<span class="${e.ok ? 'cm-ok' : 'cm-fora'}">Tom ${ordemDe(e.k)} (ref ${e.refN})${e.proprio ? ' · o dela' : ''} ${e.ok ? '✔' : '✘'} ${escH(cfg.formula)} ${num(e.pior, 2)}</span>`).join(' · ') : '';
    const falta = faltamMedir();
    const prox = av.n != null ? proximoNumero(av.n) : null;
    const pode = con && livre && !concluida;
    return `<div class="card cm-os"><div class="cm-os-grade"><div class="cm-os-esq">
      <div class="cm-botoes cm-entre">
        <div><span class="cm-fraco">Avaliando as amostras da</span> <b>${escH(S.os.rotulo)}</b></div>
        <div class="cm-botoes">
          <button type="button" class="btn small ${S.janela === 'rel' ? 'primary' : ''}" onclick="Colorimetro.alternar('rel')">Relatório de tonalidades</button>
          <button type="button" class="btn small ghost" onclick="Colorimetro.avLimpar()"
            title="Apaga a avaliação desta OS (todas as fases) e recomeça na Amostra 1">Limpar avaliação</button>
          <button type="button" class="btn small ghost" onclick="Colorimetro.encerrarOS()">Medir sem OS</button></div></div>
      ${htmlFases()}
      ${concluida ? `<div class="cm-aviso" style="background:var(--ok-bg);color:var(--ok)">${escH(fase.rotulo)} concluída. Para repetir esta fase, escolha-a em "Reavaliar a fase".</div>` : `
      <div class="cm-av">
        ${rib ? htmlRibanaPorTom(av, rib, fase) : `<div class="cm-av-bloco"><div class="cm-av-rot">${S.fases.length > 1 ? escH(fase.rotulo) + ' · ' : ''}Tons pelo rolo de referência</div>
          <div class="cm-peq">${tons.length ? tons.map((t, i) => `<b>Tom ${i + 1}</b> (ref ${t.refN}): ${t.membros.join(', ')}`).join('<br>') : 'nenhuma amostra medida ainda'}
          ${numerosDasAmostras().filter(n => tomDe(n) === 0).length ? `<br><span class="cm-fora">Não usar: ${numerosDasAmostras().filter(n => tomDe(n) === 0).join(', ')}</span>` : ''}</div></div>`}
        <div class="cm-av-bloco"><div class="cm-av-rot">Medidas nesta fase</div>
          <div class="cm-peq">${numerosDasAmostras().length}${totalDe() ? ' de ' + totalDe() : ''} ${numerosDasAmostras().length === 1 ? 'amostra' : 'amostras'}
          ${totalDe() ? `<br>${falta.length ? `Faltam medir: ${falta.join(', ')}` : '<span class="cm-ok">Todas medidas</span>'}` : ''}
          <br><span class="cm-fraco">${rib ? `Cada amostra é comparada com o rolo de referência de cada tom do ${escH(nomeCorpo)}.` : 'Cada amostra é comparada com o rolo de referência do tom (a menor amostra que ainda não tinha tom). Sem repescagem.'}</span></div></div>
      </div>
      <div class="cm-botoes">
        <button type="button" class="btn accent" onclick="Colorimetro.medir('amostra')" ${pode && av.n != null ? '' : 'disabled'}
          title="Mede a amostra atual; a média dela vira referência para todas as amostras da fase. O botão do aparelho e o F6 fazem o mesmo.">Medir amostra ${av.n != null ? av.n : ''} como referência (F5)</button>
        <button type="button" class="btn primary" onclick="Colorimetro.avTrocar()" ${av.n != null ? '' : 'disabled'}
          title="Passa para a amostra seguinte. O tom desta já foi decidido pela cor medida.">Trocar → ${prox != null ? 'amostra ' + prox : 'fim das amostras'}</button>
        <label class="cm-campo" style="max-width:120px"><span>Ir para a amostra nº</span>
          <input type="number" min="1" ${totalDe() ? `max="${totalDe()}"` : ''} value="${av.n != null ? av.n : ''}" onchange="Colorimetro.avIrPara(this.value)"></label>
        <label class="cm-campo" style="max-width:130px" title="Quantas amostras esta fase tem no total. A troca termina sozinha na última; vazio = sem limite."><span>Total de amostras</span>
          <input type="number" min="1" placeholder="sem limite" value="${totalDe() || ''}" onchange="Colorimetro.avDefinirTotal(this.value)"></label>
        <button type="button" class="btn ${av.n == null ? 'primary' : ''}" onclick="Colorimetro.avConcluirFase()" ${numerosDasAmostras().length ? '' : 'disabled'}
          title="Fecha esta fase e passa para a seguinte do enfesto">${ultimaFase ? 'Concluir a fase' : 'Concluir a fase → próxima'}</button>
      </div>`}
      ${aparelho}</div>
      <div class="cm-os-dir">
        ${concluida ? '' : `<div class="cm-av-bloco cm-av-atual cm-av-topo"><div class="cm-av-rot">Amostra atual</div>
          <div class="cm-av-num cm-av-num-grande">${av.n != null ? 'Nº ' + av.n : '—'}${av.n != null && totalDe() ? ` <span class="cm-fraco" style="font-size:.5em">de ${totalDe()}</span>` : ''}</div>
          <div class="cm-peq">${av.n != null ? `${ms.length} ${ms.length === 1 ? 'medição' : 'medições'} · ${situacao}${encaixe ? '<br>' + encaixe : ''}` : escH(avDicaFase())}</div></div>`}
        ${comparacao}</div></div></div>`;
  }

  /* A TABELA DE COMPATIBILIDADE (12/11/2026, Junior: "mostre a tabela de
     compatibilidade no relatório"). É a tabela de onde os tons saem — desde a
     opção A (rolo de referência), cada amostra contra o ROLO DE REFERÊNCIA de
     cada tom: ✔ alguma medição dela aprovada contra a referência, ✘ nenhuma,
     pela tolerância das Configurações e pela diferença real. A coluna do tom
     da própria amostra fica sombreada; "ref" marca a referência. O dE da
     melhor medição fica só na dica ao passar o mouse. */
  function htmlTabelaCompat(av, tons) {
    const nums = numerosDasAmostras(av).filter(n => (av.medidas[n] || []).length);
    if (!tons.length || nums.length < 2) return '';
    const cab = tons.map((t, i) => `<th class="cm-tc-n" title="Tom ${i + 1} · referência ${escH(nomeAm(t.refN))}">Tom ${i + 1}<br><small>ref ${t.refN}</small></th>`).join('');
    const linhas = nums.map(n => {
      const meu = tons.findIndex(t => t.k === tomDe(n, av));
      const cel = tons.map((t, i) => {
        if (t.refN === n) return '<td class="cm-tc-mesmo cm-tc-eu">ref</td>';
        const m = melhorContraRef(n, t.refN, av);
        if (!m) return '<td></td>';
        const dica = `${nomeAm(n)} × referência do Tom ${i + 1} (${nomeAm(t.refN)}), melhor medição: ${m.aprovado ? 'aprovada' : 'reprovada'} · ${cfg.formula} ${num(m.r.des[cfg.formula], 2)}`
          + ` · dL* ${num(m.r.d1[0], 2)} da* ${num(m.r.d1[1], 2)} db* ${num(m.r.d1[2], 2)}`;
        return `<td class="${i === meu ? 'cm-tc-mesmo ' : ''}${m.aprovado ? 'cm-ok' : 'cm-fora'}" title="${escH(dica)}">${m.aprovado ? '✔' : '✘'}</td>`;
      }).join('');
      const rot = tomDe(n, av) === 0 ? 'não usar' : meu >= 0 ? 'Tom ' + (meu + 1) : '—';
      return `<tr><th class="cm-tc-n" title="${escH(nomeAm(n))} · ${escH(rot)}">${n}</th><th class="cm-tc-tom">${escH(rot)}</th>${cel}</tr>`;
    }).join('');
    return `<details class="cm-tc" open><summary class="cm-sub" style="cursor:pointer">Tabela de compatibilidade <span class="cm-fraco">(cada amostra × rolo de referência de cada tom · ✔ aprovada · ✘ reprovada · sombreado = o tom dela)</span></summary>
      <div class="cm-rolar"><table class="cm-tc-tab"><thead><tr><th>Amostra</th><th></th>${cab}</tr></thead><tbody>${linhas}</tbody></table></div></details>`;
  }

  // O relatório da ribana comparada com o corpo: a quais tons do corpo cada uma serve.
  function htmlRelFaseRibana(f, av, rib) {
    const nc = escH(rib.fCorpo.nome);
    const tonsCorpo = tonsDoRelatorio(rib.corpo);
    const nums = numerosDasAmostras(av);
    const linhas = nums.map(n => {
      const fixo = av.fixo && (n in av.fixo) ? av.fixo[n] : undefined;
      const auto = fixo === undefined;
      const rc = ribanaContraCorpo(n, av, rib);
      const serve = rc ? rc.serve : [];
      const opc = [`<option value="auto" ${auto ? 'selected' : ''}>Automático${auto ? (serve.length ? ' · ' + serve.map(o => 'Tom ' + o).join(', ') : ' · incompatível') : ''}</option>`]
        .concat(tonsCorpo.map((t, i) => `<option value="${i + 1}" ${fixo === i + 1 ? 'selected' : ''}>Tom ${i + 1} do ${nc} (fixo)</option>`))
        .concat([`<option value="0" ${fixo === 0 ? 'selected' : ''}>Não usar</option>`]);
      const cor = corDaAmostra(n, av);
      const casaTxt = rc ? rc.tons.map(t => `<span class="${t.ok ? 'cm-ok' : 'cm-fora'}">Tom ${t.ordem} ${t.ok ? '✔' : '✘'}</span>`).join(' ') : '';
      return `<tr><td>${cor ? `<span class="cm-amostra-cor" style="background:${labHex(...cor)}"></span>` : ''}${escH(nomeAm(n))}</td>
        <td class="cm-n">${(av.medidas[n] || []).length}</td><td>${casaTxt}</td>
        <td><select onchange="Colorimetro.relMudar(${f.i}, ${n}, this.value)">${opc.join('')}</select></td></tr>`;
    }).join('');
    const resumo = tonsCorpo.map((t, i) => {
      const servem = nums.filter(n => tomDe(n, av) !== 0 && (av.serve[n] || []).includes(i + 1));
      return `<div class="cm-tom ${servem.length ? '' : 'cm-tom-ruim'}"><b>${nc} Tom ${i + 1}</b> (${t.membros.join(', ')}) — ${servem.length ? 'ribana ' + servem.join(', ') : '<span class="cm-fora">nenhuma ribana compatível</span>'}</div>`;
    }).join('');
    const incomp = nums.filter(n => ehIncompativelCorpo(n, av));
    // A tabela: cada ribana × o rolo de referência de cada tom do corpo.
    const cabTom = '';
    const cab = tonsCorpo.map((t, i) => `<th class="cm-tc-n" title="${nc} Tom ${i + 1} · referência ${t.refN}">Tom ${i + 1}<br><small>ref ${t.refN}</small></th>`).join('');
    const tab = nums.filter(n => (av.medidas[n] || []).length).map(n => {
      return `<tr><th class="cm-tc-n">${n}</th>${tonsCorpo.map((t, i) => {
        const m = melhorContra(av.medidas[n], labRef(t.refN, rib.corpo));
        if (!m) return '<td></td>';
        const dica = `Ribana ${n} × referência do Tom ${i + 1} do ${rib.fCorpo.nome} (amostra ${t.refN}), melhor medição: ${m.aprovado ? 'aprovada' : 'reprovada'} · ${cfg.formula} ${num(m.r.des[cfg.formula], 2)} · dL* ${num(m.r.d1[0], 2)} da* ${num(m.r.d1[1], 2)} db* ${num(m.r.d1[2], 2)}`;
        return `<td class="${m.aprovado ? 'cm-ok' : 'cm-fora'}" title="${escH(dica)}">${m.aprovado ? '✔' : '✘'}</td>`;
      }).join('')}</tr>`;
    }).join('');
    const marcas = marcasDaFase(f);
    return `<div class="cm-rel-fase">
      <div class="cm-sub">${escH(f.rotulo)} <span class="cm-fraco">(comparada com o ${nc}${marcas.length ? ', ' + escH(marcas.join(', ')) : ''})</span></div>
      <div class="cm-rolar"><table class="cm-lista"><thead><tr><th>Amostra</th><th>Medições</th><th>Serve ao ${nc}</th><th>Tom</th></tr></thead>
        <tbody>${linhas}</tbody></table></div>
      ${resumo || '<div class="cm-fraco">O corpo ainda não tem tom.</div>'}
      ${incomp.length ? `<div class="cm-fora" style="margin-top:6px">Incompatíveis com o ${nc}: ${incomp.map(nomeAm).join(', ')}</div>` : ''}
      ${tab ? `<details class="cm-tc" open><summary class="cm-sub" style="cursor:pointer">Tabela de compatibilidade · ribana × ${nc} <span class="cm-fraco">(cada ribana × rolo de referência de cada tom do ${nc} · ✔ aprovada · ✘ reprovada)</span></summary>
        <div class="cm-rolar"><table class="cm-tc-tab"><thead><tr><th>Ribana</th>${cab}</tr></thead><tbody>${tab}</tbody></table></div></details>` : ''}
    </div>`;
  }

  // O relatório de uma fase: amostras, tons e o seletor para corrigir.
  function htmlRelFase(f) {
    const av = S.avOS.fases[f.i], fo = cfg.formula;
    if (av.contraCorpo && corpoDaRibana(av)) return htmlRelFaseRibana(f, av, corpoDaRibana(av));
    const nums = numerosDasAmostras(av);
    const tons = tonsDoRelatorio(av);
    const ordem = {};
    tons.forEach((t, i) => { ordem[t.k] = i + 1; });
    const ks = tons.map(t => t.k);
    // Contra o rolo de referência do tom dela: o dE da melhor medição.
    const piorDe = {}, ehRefDe = {};
    tons.forEach(t => t.membros.forEach(n => {
      if (n === t.refN) { ehRefDe[n] = true; return; }
      const m = melhorContraRef(n, t.refN, av);
      piorDe[n] = m ? m.d : null;
    }));
    const linhas = nums.map(n => {
      const t = tomDe(n, av);
      const fixo = av.fixo && (n in av.fixo) ? av.fixo[n] : undefined;
      const todas = av.medidas[n] || [];
      const auto = fixo === undefined;
      const opc = [`<option value="auto" ${auto ? 'selected' : ''}>Automático${auto && ordem[t] ? ' · Tom ' + ordem[t] : ''}</option>`]
        .concat(ks.map(k => `<option value="${k}" ${!auto && fixo === k ? 'selected' : ''}>Tom ${ordem[k]} (fixo)</option>`))
        .concat([`<option value="novo">Tom ${tons.length + 1} (novo, só esta amostra)</option>`,
          `<option value="0" ${fixo === 0 ? 'selected' : ''}>Não usar</option>`]);
      const cor = corDaAmostra(n, av);
      return `<tr><td>${cor ? `<span class="cm-amostra-cor" style="background:${labHex(...cor)}"></span>` : ''}${escH(nomeAm(n))}</td>
        <td class="cm-n">${todas.length}</td>
        <td class="cm-n">${ehRefDe[n] ? 'referência' : piorDe[n] != null ? num(piorDe[n], 2) : ''}</td>
        <td><select onchange="Colorimetro.relMudar(${f.i}, ${n}, this.value)">${opc.join('')}</select></td></tr>`;
    }).join('');
    const resumo = tons.map((t, i) => `<div class="cm-tom ${t.ruins.length ? 'cm-tom-ruim' : ''}">
        <b>Tom ${i + 1}</b> — referência ${escH(nomeAm(t.refN))} · ${t.membros.map(nomeAm).join(', ')}
        <span class="cm-fraco">${t.membros.length > 1 ? `· ${escH(fo)} máx contra a referência ${num(t.pior, 2)}` : '· só ela'}</span>
        ${t.ruins.length ? `<div class="cm-fora">Marcadas à mão e reprovadas contra a referência: ${t.ruins.map(([n, d]) => `${escH(nomeAm(n))} (${num(d, 2)})`).join('; ')}</div>` : ''}
      </div>`).join('');
    const pend = avPendentes(av);
    const marcas = marcasDaFase(f);
    if (f.corDe >= 0 && S.fases[f.corDe]) marcas.push(`no recado vai com ${S.fases[f.corDe].nome}, tom a tom`);
    return `<div class="cm-rel-fase">
      <div class="cm-sub">${escH(S.fases.length > 1 ? f.rotulo : (f.tecido || 'Amostras'))}${marcas.length ? ` <span class="cm-fraco">(${escH(marcas.join(', '))})</span>` : ''}</div>
      <div class="cm-rolar"><table class="cm-lista"><thead><tr><th>Amostra</th><th>Medições</th><th title="Diferença da melhor medição contra o rolo de referência do tom">${escH(fo)} contra a referência</th><th>Tom</th></tr></thead>
        <tbody>${linhas}</tbody></table></div>
      ${resumo || '<div class="cm-fraco">Nenhuma amostra num tom.</div>'}
      ${htmlTabelaCompat(av, tons)}
      ${pend.length ? `<div class="cm-fora" style="margin-top:6px">A reavaliar: ${pend.map(nomeAm).join(', ')}</div>` : ''}
    </div>`;
  }

  function htmlRel() {
    if (S.janela !== 'rel') return '';
    const topo = `<div class="cm-janela-topo"><div class="card-title">Relatório de tonalidades · ${escH(S.os ? S.os.rotulo : '')}</div>${botaoFechar()}</div>`;
    if (!S.avOS) return `<div class="card cm-rel">${topo}<div class="cm-fraco">Escolha a OS pelo "⋯ → avaliar amostras" da lista de OS.</div></div>`;
    const fo = cfg.formula, tol = tolAtual();
    const fases = fasesComMedicao();
    if (!fases.length) return `<div class="card cm-rel">${topo}<div class="cm-fraco">Ainda não há amostras medidas para a ${escH(S.os.nome)}.</div></div>`;
    const ehOP = S.os.doc === 'OP';
    const faltam = S.fases.filter(f => !faseFeita(S.avOS, f.i));
    const algumTom = fases.some(f => tonsDoRelatorio(S.avOS.fases[f.i]).length);
    const salvo = S.relSalvo && S.relSalvo.osId === S.os.id
      ? `<span class="cm-ok">✔ salvo ${ehOP ? 'na folha da OP' : 'nas observações'} às ${horaTxt(S.relSalvo.em).slice(0, 5)}</span>
         <button type="button" class="btn small ghost" onclick="${ehOP ? 'verOP' : 'verOS'}('${escH(S.os.id)}')">Ver a folha da ${ehOP ? 'OP' : 'OS'}</button>` : '';
    return `<div class="card cm-rel">${topo}
      <div class="cm-fraco" style="margin-bottom:8px">Por fase de enfesto e tecido. Cada tom tem um rolo de referência — a menor amostra que ainda não tinha tom —, e entra nele toda amostra com alguma medição aprovada contra a referência
        (<b>${escH(fo)}</b> até <b>${numTol(tol)}</b> e dL*, da*, db* dentro, pela diferença real, ${escH(cfg.iluminante1)}°). Sem repescagem: cada medição nova refaz a classificação. Fixe o tom de qualquer amostra aqui; o texto abaixo acompanha.</div>
      ${fases.map(htmlRelFase).join('')}
      ${faltam.length && S.fases.length > 1 ? `<div class="cm-fora" style="margin-top:8px">Fases ainda não concluídas: ${faltam.map(f => escH(f.rotulo)).join('; ')}</div>` : ''}
      <div class="cm-sub">${ehOP ? `Vai para a folha da ${escH(S.os.nome)}` : `Vai para as observações da OS ${escH(S.os.numero)}`}</div>
      <div class="cm-previa">${escH(algumTom ? textoRelatorio() : '—')}</div>
      <div class="cm-botoes">
        <button type="button" class="btn primary" onclick="Colorimetro.relSalvar()" ${algumTom ? '' : 'disabled'}>${ehOP ? 'Salvar na folha da OP' : 'Salvar nas observações da OS'}</button>
        <button type="button" class="btn small ghost" onclick="Colorimetro.exportarAvaliacao()"
          title="Baixa um CSV com TODAS as medições guardadas nesta avaliação, de todas as fases e dias: fase, amostra, horário, L*a*b* e o tom atual">Exportar dados da avaliação</button>
        ${salvo}
      </div></div>`;
  }

  function abrirRelatorio() { S.janela = 'rel'; desenhar(); rolarParaJanela(); }

  /* ---------------- configurações ---------------- */
  function aplicarCfg() {
    const v = id => document.getElementById(id);
    const erros = [];
    const numero = (id, rot, min) => {
      const x = lerNum(v(id).value);
      if (x == null || (min != null && x < min)) { erros.push(rot); return null; }
      return x;
    };
    const novo = JSON.parse(JSON.stringify(cfg));
    novo.iluminante1 = v('cm-il1').value;
    novo.iluminante2 = v('cm-il2').value;
    novo.modo = v('cm-modo').value;
    novo.formula = v('cm-formula').value;
    novo.media = Math.max(1, Math.round(numero('cm-media', 'Média', 1) || 1));
    novo.nome_padrao = v('cm-nome-p').value.trim() || 'Padrão';
    novo.nome_amostra = v('cm-nome-a').value.trim() || 'Amostra';
    novo.cmc_l = numero('cm-cmc-l', 'CMC l', 0.01);
    novo.cmc_c = numero('cm-cmc-c', 'CMC c', 0.01);
    novo.k94 = [0, 1, 2].map(i => numero('cm-k94-' + i, 'k do dE94', 0.01));
    novo.k00 = [0, 1, 2].map(i => numero('cm-k00-' + i, 'k do dE00', 0.01));
    Object.keys(novo.tol).forEach(k => { novo.tol[k] = numero('cm-tol-' + k, 'tolerância ' + k, 0); });
    if (erros.length) { S.aviso = 'Confira: ' + [...new Set(erros)].join(', ') + '.'; desenhar(); return; }
    cfg = novo;
    salvarCfg();
    S.leituras.forEach(recalcular);
    S.janela = null;
    S.aviso = '';
    status('Configurações aplicadas');
  }
  function resetarCfg() {
    perguntar({ titulo: 'Voltar ao padrão', texto: 'Voltar todas as configurações (tolerâncias inclusive) aos valores do programa?',
      sim: 'Voltar ao padrão', executar: resetarCfgAgora });
  }
  function resetarCfgAgora() {
    const ultimo = cfg.ultimo_aparelho;
    cfg = JSON.parse(JSON.stringify(CFG_PADRAO));
    cfg.ultimo_aparelho = ultimo;
    salvarCfg();
    S.leituras.forEach(recalcular);
    desenhar();
  }

  /* ---------------- tela ---------------- */
  const painel = () => document.getElementById('colorimetro-painel');
  const visivel = () => {
    const sec = document.querySelector('section.page[data-page="colorimetro"]');
    return !!(sec && !sec.classList.contains('hidden'));
  };

  function tabelaPagina(ref, amo, r) {
    const il1 = cfg.iluminante1, il2 = cfg.iluminante2, tol = cfg.tol;
    const f = (v, sinal) => v == null ? '-' : ((sinal && v > 0 ? '+' : '') + num(v));
    const ct = (v, k) => v == null ? '' : (dentro(v, tol[k]) ? 'cm-ok' : 'cm-fora');
    const vazio3 = [null, null, null];
    const P = S.os && S.av ? 'Referência' : 'Padrão';
    let titulo, cab, linhas;
    if (S.pagina === 0) {
      titulo = `L*a*b* · ${il1}°`; cab = ['L*', 'a*', 'b*'];
      const d = r ? r.d1 : vazio3;
      linhas = [[P, (ref ? ref.lab1 : vazio3).map(v => f(v))], ['Amostra', (amo ? amo.lab1 : vazio3).map(v => f(v))],
        ['Diferença', d.map(v => v == null ? '-' : ((v > 0 ? '+' : '') + num(v, 2)))]];
      if (r) linhas.push(['', r.dicas.map(x => x[0]), r.dicas.map(x => x[1] ? 'cm-ok' : 'cm-fora')]);
    } else if (S.pagina === 1) {
      titulo = `L*C*h° · ${il1}°`; cab = ['L*', 'C*', 'h°'];
      const lc = x => x ? [x.lab1[0], ...lch(...x.lab1)] : vazio3;
      const d = r ? r.d : [null, null, null, null, null];
      linhas = [[P, lc(ref).map(v => f(v))], ['Amostra', lc(amo).map(v => f(v))],
        ['Diferença', [f(d[0], true), f(d[3], true), r ? 'ΔH ' + f(d[4], true) : '-'], [ct(d[0], 'dL'), ct(d[3], 'dC'), ct(d[4], 'dH')]]];
    } else if (S.pagina === 2) {
      titulo = `Diferença de cor · ${il1}°`; cab = ['ΔE', 'Tolerância', 'Resultado'];
      linhas = FORMULAS.map(fo => {
        const v = r ? r.des[fo] : null;
        const ok = v == null ? null : v <= tol[fo];
        return [fo, [f(v), numTol(tol[fo]), ok == null ? '-' : (ok ? 'OK' : 'Fora')], [ct(v, fo), 'cm-fraco', ok == null ? '' : (ok ? 'cm-ok' : 'cm-fora')]];
      });
    } else {
      titulo = `Segundo iluminante · ${il2}° (metamerismo)`; cab = ['L*', 'a*', 'b*'];
      const a = ref && ref.lab2, b = amo && amo.lab2;
      const dif = a && b ? [0, 1, 2].map(i => b[i] - a[i]) : vazio3;
      linhas = [[P, (a || vazio3).map(v => f(v))], ['Amostra', (b || vazio3).map(v => f(v))],
        ['Diferença', dif.map(v => f(v, true)), dif.map((v, i) => ct(v, ['dL', 'da', 'db'][i]))]];
      if (a && b) linhas.push(['ΔE*ab', [num(deAb(a, b)), '', '']]);
    }
    const abas = ['L*a*b*', 'L*C*h°', 'ΔE', '2º iluminante'].map((t, i) =>
      `<button type="button" class="btn small ${S.pagina === i ? 'primary' : 'ghost'}" onclick="Colorimetro.pagina(${i})">${t}</button>`).join(' ');
    return `<div class="cm-abas">${abas}</div>
      <div class="cm-tit">${escH(titulo)}</div>
      <table class="cm-tab"><thead><tr><th></th>${cab.map(c => `<th>${escH(c)}</th>`).join('')}</tr></thead>
      <tbody>${linhas.map(([rot, vs, cores]) => `<tr><td class="cm-rot">${escH(rot)}</td>${vs.map((x, j) =>
        `<td class="${(cores && cores[j]) || ''}">${escH(x)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  }

  function blocoCor(l, rotulo, tipo) {
    const cor = l && l.lab1 ? labHex(...l.lab1) : '';
    const claro = l && l.lab1 && l.lab1[0] > 72;
    const acao = S.dev && !S.ocupado ? `onclick="Colorimetro.medir('${tipo}')" title="Medir ${rotulo.toLowerCase()}"` : '';
    return `<div class="cm-bloco ${cor ? '' : 'cm-bloco-vazio'}" style="${cor ? `background:${cor};color:${claro ? '#1d2230' : '#fff'}` : ''}" ${acao}>
      <span>${escH(l ? l.nome : rotulo)}</span></div>`;
  }

  /* O VISOR DA AVALIAÇÃO DE OS (12/11/2026, Junior: "faça o visor mostrar se é
     compatível ou não com cada referência medida no momento da medição").
     A faixa única "Aprovado/Reprovado" comparava a ÚLTIMA leitura com a primeira
     leitura de UMA amostra (sobra do modelo de rodadas) e podia até comparar a
     amostra com ela mesma. Agora o visor mostra a amostra atual — pela média,
     como a classificação — contra CADA amostra já medida da fase, uma a uma:
     ✔ compatível / ✘ incompatível pela tolerância das Configurações. A faixa de
     cima resume: o tom em que ela ficou e com quantas é compatível. O dE vai só
     na dica do mouse. Fora da OS, o visor de padrão × amostra segue igual. */
  // Visor da ribana: ela contra CADA amostra do corpo, agrupadas pelo tom do corpo.
  function htmlVisorRibana(av, n, cor, nMed, rib) {
    const rc = ribanaContraCorpo(n, av, rib);
    const nc = escH(rib.fCorpo.nome);
    const naoUsar = tomDe(n, av) === 0;
    const faixa = naoUsar ? `<div class="cm-de cm-de-fora">Amostra ${n} · não usar</div>`
      : rc.serve.length
        ? `<div class="cm-de cm-de-ok">Amostra ${n} → ${rc.serve.map(o => 'Tom ' + o).join(', ')} do ${nc}</div>
           <div class="cm-fraco cm-centro">Aprovada contra o rolo de referência ${rc.serve.length > 1 ? 'desses tons' : 'desse tom'} do ${nc}</div>`
        : `<div class="cm-de cm-de-fora">Amostra ${n} → incompatível com o ${nc}</div>
           <div class="cm-fraco cm-centro">Reprovada contra o rolo de referência de todos os tons do ${nc}</div>`;
    const grupos = rc.tons.map(t => {
      const chips = rc.porAmostra.filter(e => t.membros.includes(e.x)).map(e => {
        const dica = `${nomeAm(n)} × referência ${e.x} do ${rib.fCorpo.nome}, melhor medição: ${e.ok ? 'aprovada' : 'reprovada'} · ${cfg.formula} ${num(e.r.des[cfg.formula], 2)}`
          + ` · dL* ${num(e.r.d1[0], 2)} da* ${num(e.r.d1[1], 2)} db* ${num(e.r.d1[2], 2)}` + (e.r.desc.length ? ` (${e.r.desc.join(', ')})` : '');
        return `<div class="cm-ref ${e.ok ? 'cm-ref-ok' : 'cm-ref-fora'}" title="${escH(dica)}">
          <span class="cm-ref-cor" style="background:${labHex(...(labRef(e.x, rib.corpo) || corDaAmostra(e.x, rib.corpo)))}"></span>ref ${e.x}<b>${e.ok ? '✔' : '✘'}</b></div>`;
      }).join('');
      return `<div class="cm-peq" style="margin:4px 0 2px"><b class="${t.ok ? 'cm-ok' : 'cm-fora'}">${nc} Tom ${t.ordem} ${t.ok ? '✔' : '✘'}</b></div><div class="cm-refs">${chips}</div>`;
    }).join('');
    const claro = cor[0] > 72;
    return `<div class="cm-blocos" style="grid-template-columns:1fr">
        <div class="cm-bloco" style="background:${labHex(...cor)};color:${claro ? '#1d2230' : '#fff'};cursor:default"><span>${escH(nomeAm(n))} · média de ${nMed} ${nMed === 1 ? 'medição' : 'medições'}</span></div></div>
      ${faixa}
      <div class="cm-sub" style="margin:6px 0 2px">Contra o rolo de referência de cada tom do ${nc}</div>${grupos || '<div class="cm-fraco">O corpo ainda não tem tom.</div>'}
      <div class="cm-fraco cm-centro cm-peq" style="margin-top:6px">L* ${num(cor[0])} · a* ${num(cor[1])} · b* ${num(cor[2])} · ${escH(cfg.formula)} até ${numTol(tolAtual())} · ${escH(cfg.iluminante1)}°</div>`;
  }

  function htmlVisorOS() {
    const av = S.av, n = av && av.n;
    const cor = n != null ? corDaAmostra(n) : null;
    if (!cor) {
      return `<div class="cm-de">—</div><div class="cm-fraco cm-centro">${n != null
        ? `Meça a Amostra ${n}: ela será comparada com o rolo de referência de cada tom.` : escH(avDicaFase())}</div>`;
    }
    const nMed = (av.medidas[n] || []).length;
    const rib = av.contraCorpo ? corpoDaRibana(av) : null;
    if (rib) return htmlVisorRibana(av, n, cor, nMed, rib);
    // O ROLO DE REFERÊNCIA (12/11/2026): a amostra atual contra a referência de
    // cada tom — a mesma conta que decide o tom dela.
    const tons = tonsDoRelatorio();
    const res = tons.map((t, i) => {
      if (t.refN === n) return { t, ordem: i + 1, eu: true };
      const m = melhorContraRef(n, t.refN);
      return { t, ordem: i + 1, eu: false, ok: !!(m && m.aprovado), r: m ? m.r : null };
    });
    const tAqui = tons.find(t => t.k === tomDe(n));
    const ordem = tAqui ? tons.indexOf(tAqui) + 1 : 0;
    const anteriores = tons.filter(t => t.refN < n).map(t => t.refN);
    let faixa;
    if (tomDe(n) === 0) faixa = `<div class="cm-de cm-de-fora">Amostra ${n} · não usar</div>`;
    else if (tAqui && tAqui.refN === n) faixa = anteriores.length
      ? `<div class="cm-de cm-de-fora">Amostra ${n} → referência do Tom ${ordem}</div>
         <div class="cm-fraco cm-centro">Reprovada contra a referência de ${anteriores.length > 1 ? 'todos os tons anteriores' : 'Tom 1'}: abre um tom novo, e as próximas são comparadas com ela</div>`
      : `<div class="cm-de">Amostra ${n} · referência do Tom ${ordem}</div><div class="cm-fraco cm-centro">As próximas amostras são comparadas com este rolo.</div>`;
    else faixa = `<div class="cm-de cm-de-ok">Amostra ${n} → Tom ${ordem}</div>
         <div class="cm-fraco cm-centro">Aprovada contra o rolo de referência (Amostra ${tAqui ? tAqui.refN : '?'})</div>`;
    const chips = res.map(e => {
      const corRef = labRef(e.t.refN);
      if (e.eu) return `<div class="cm-ref" title="Esta amostra é a referência do Tom ${e.ordem}"><span class="cm-ref-cor" style="background:${labHex(...(corRef || cor))}"></span>Tom ${e.ordem} · é a referência</div>`;
      const dica = e.r ? `${nomeAm(n)} × referência do Tom ${e.ordem} (${nomeAm(e.t.refN)}), melhor medição: ${e.ok ? 'aprovada' : 'reprovada'} · ${cfg.formula} ${num(e.r.des[cfg.formula], 2)}`
        + ` · dL* ${num(e.r.d1[0], 2)} da* ${num(e.r.d1[1], 2)} db* ${num(e.r.d1[2], 2)}` + (e.r.desc.length ? ` (${e.r.desc.join(', ')})` : '') : '';
      return `<div class="cm-ref ${e.ok ? 'cm-ref-ok' : 'cm-ref-fora'}" title="${escH(dica)}">
        <span class="cm-ref-cor" style="background:${labHex(...(corRef || cor))}"></span>Tom ${e.ordem} · ref ${e.t.refN}<b>${e.ok ? '✔' : '✘'}</b></div>`;
    }).join('');
    const claro = cor[0] > 72;
    return `<div class="cm-blocos" style="grid-template-columns:1fr">
        <div class="cm-bloco" style="background:${labHex(...cor)};color:${claro ? '#1d2230' : '#fff'};cursor:default"><span>${escH(nomeAm(n))} · ${nMed} ${nMed === 1 ? 'medição' : 'medições'}</span></div></div>
      ${faixa}
      ${res.length ? `<div class="cm-sub" style="margin:6px 0 4px">Contra o rolo de referência de cada tom</div><div class="cm-refs">${chips}</div>` : ''}
      <div class="cm-fraco cm-centro cm-peq" style="margin-top:6px">L* ${num(cor[0])} · a* ${num(cor[1])} · b* ${num(cor[2])} (média) · ${escH(cfg.formula)} até ${numTol(tolAtual())} · ${escH(cfg.iluminante1)}°</div>`;
  }

  function linhaLista(l) {
    const [C, h] = lch(...l.lab1);
    let dE = '', dE00 = '', res = '', cls = '';
    const ref = refDe(l);
    if (l.tipo === 'amostra' && ref) {
      const r = comparar(ref, l);
      dE = num(r.des['dE*ab']); dE00 = num(r.des.dE00);
      res = r.aprovado ? 'Aprovado' : 'Reprovado';
      cls = r.aprovado ? 'cm-ok' : 'cm-fora';
    } else if (l.tipo === 'padrao') res = l === S.padrao ? '(padrão)' : 'padrão antigo';
    return `<tr class="${l.tipo === 'padrao' ? 'cm-lin-padrao' : ''} ${l === S.atual ? 'cm-lin-atual' : ''}" onclick="Colorimetro.selecionar('${l.id}')">
      <td>${escH(l.nome)}</td><td>${horaTxt(l.horario)}</td>${l.lab1.map(v => `<td class="cm-n">${num(v)}</td>`).join('')}
      <td class="cm-n">${num(C)}</td><td class="cm-n">${num(h)}</td><td class="cm-n">${dE}</td><td class="cm-n">${dE00}</td>
      <td class="${cls}">${res}</td>
      <td class="cm-acoes" onclick="event.stopPropagation()">
        <button type="button" class="btn small ghost" title="Usar como padrão" onclick="Colorimetro.usarPadrao('${l.id}')">★</button>
        <button type="button" class="btn small ghost" title="Renomear" onclick="Colorimetro.renomear('${l.id}')">✎</button>
        <button type="button" class="btn small ghost" title="Tirar da lista (o histórico guarda)" onclick="Colorimetro.excluir('${l.id}')">✕</button>
      </td></tr>`;
  }

  function htmlCfg() {
    if (S.janela !== 'cfg') return '';
    const ils = T ? Object.keys(T.illuminants) : [cfg.iluminante1, cfg.iluminante2];
    const opt = (lista, atual) => lista.map(x => `<option value="${escH(x[0])}" ${String(atual) === String(x[0]) ? 'selected' : ''}>${escH(x[1])}</option>`).join('');
    const campo = (id, rot, val, larg) => `<label class="cm-campo" style="${larg ? 'max-width:' + larg : ''}"><span>${rot}</span><input id="${id}" value="${escH(numTol(val))}" inputmode="decimal"></label>`;
    const tolRot = { dL: 'dL* ±', da: 'da* ±', db: 'db* ±', 'dE*ab': 'dE*ab ≤', dC: 'dC* ±', dH: 'dH* ±', "dL'": "dL' ±", "dC'": "dC' ±", "dH'": "dH' ±", dE00: 'dE00 ≤', dEcmc: 'dEcmc ≤', dE94: 'dE94 ≤' };
    return `<div class="card cm-cfg">
      <div class="cm-janela-topo"><div class="card-title">Configurações do colorímetro <span class="cm-fraco">(valem neste computador)</span></div>${botaoFechar()}</div>
      <div class="cm-grade">
        <label class="cm-campo"><span>Iluminante principal</span><select id="cm-il1">${opt(ils.map(x => [x, x + '°']), cfg.iluminante1)}</select></label>
        <label class="cm-campo"><span>Segundo iluminante (metamerismo)</span><select id="cm-il2">${opt(ils.map(x => [x, x + '°']), cfg.iluminante2)}</select></label>
        <label class="cm-campo"><span>Modo de medição</span><select id="cm-modo">${opt([['auto', 'Automático (o do aparelho)'], ...Object.entries(MODOS)], cfg.modo)}</select></label>
        <label class="cm-campo"><span>Aprovação por</span><select id="cm-formula">${opt(FORMULAS.map(x => [x, x]), cfg.formula)}</select></label>
        ${campo('cm-media', 'Média de quantas leituras', cfg.media)}
        <label class="cm-campo"><span>Nome do padrão</span><input id="cm-nome-p" value="${escH(cfg.nome_padrao)}"></label>
        <label class="cm-campo"><span>Nome da amostra</span><input id="cm-nome-a" value="${escH(cfg.nome_amostra)}"></label>
      </div>
      <div class="cm-sub">Tolerâncias</div>
      <div class="cm-grade">${Object.keys(cfg.tol).map(k => campo('cm-tol-' + k, tolRot[k] || k, cfg.tol[k])).join('')}</div>
      <div class="cm-sub">Fórmulas</div>
      <div class="cm-grade">
        ${campo('cm-cmc-l', 'CMC l', cfg.cmc_l)}${campo('cm-cmc-c', 'CMC c', cfg.cmc_c)}
        ${[0, 1, 2].map(i => campo('cm-k94-' + i, 'dE94 ' + ['kL', 'kC', 'kH'][i], cfg.k94[i])).join('')}
        ${[0, 1, 2].map(i => campo('cm-k00-' + i, 'dE00 ' + ['kL', 'kC', 'kH'][i], cfg.k00[i])).join('')}
      </div>
      <div class="cm-botoes">
        <button type="button" class="btn primary" onclick="Colorimetro.aplicarCfg()">Aplicar</button>
        <button type="button" class="btn" onclick="Colorimetro.fecharCfg()">Cancelar</button>
        <button type="button" class="btn ghost" onclick="Colorimetro.resetarCfg()">Voltar ao padrão do app</button>
      </div></div>`;
  }

  function htmlHist() {
    if (S.janela !== 'hist') return '';
    const fo = cfg.formula;
    let aprov = 0, reprov = 0;
    const porId = {};
    S.histLidas.forEach(l => { porId[l.id] = l; });
    const linhas = S.histLidas.slice().reverse().map(l => {
      let dE = '', dE00 = '', res = '', cls = '';
      const ref = porId[l.refId];
      if (ref) {
        const r = comparar(ref, l);
        dE = num(r.des['dE*ab']); dE00 = num(r.des.dE00);
        res = r.aprovado ? 'Aprovado' : 'Reprovado'; cls = r.aprovado ? 'cm-ok' : 'cm-fora';
        if (r.aprovado) aprov++; else reprov++;
      }
      return `<tr class="${l.tipo === 'padrao' ? 'cm-lin-padrao' : ''}"><td>${horaTxt(l.horario)}</td><td>${escH(l.nome)}</td>
        <td>${l.tipo === 'padrao' ? 'Padrão' : 'Amostra'}</td>${l.lab1.map(v => `<td class="cm-n">${num(v)}</td>`).join('')}
        <td>${escH(ref ? ref.nome : (l.refNome || ''))}</td><td class="cm-n">${dE}</td><td class="cm-n">${dE00}</td><td class="${cls}">${res}</td>
        <td><button type="button" class="btn small ghost" onclick="Colorimetro.histUsarPadrao('${l.id}')">Usar como padrão</button></td></tr>`;
    }).join('');
    return `<div class="card cm-hist">
      <div class="cm-janela-topo"><div class="card-title">Histórico deste computador</div>${botaoFechar()}</div>
      ${S.histCarregando ? '<div class="cm-fraco">Carregando…</div>' : S.histDias.length ? `<div class="cm-botoes">
        <label class="cm-campo" style="max-width:180px"><span>Dia</span><select onchange="Colorimetro.carregarDia(this.value)">${S.histDias.map(d =>
          `<option value="${d}" ${d === S.histDia ? 'selected' : ''}>${d.split('-').reverse().join('/')}</option>`).join('')}</select></label>
        <button type="button" class="btn" onclick="Colorimetro.histExportar()">Exportar este dia (CSV)</button>
        <button type="button" class="btn ghost" onclick="Colorimetro.fecharHist()">Fechar</button></div>
        <div class="cm-fraco" style="margin:6px 0">${S.histLidas.length} medições · ${aprov} aprovadas · ${reprov} reprovadas
          (resultado pelas regras de agora: ${fo} ≤ ${numTol(cfg.tol[fo])})</div>
        <div class="cm-rolar"><table class="cm-lista"><thead><tr><th>Hora</th><th>Nome</th><th>Tipo</th><th>L*</th><th>a*</th><th>b*</th>
          <th>Comparado com</th><th>dE*ab</th><th>dE00</th><th>Resultado</th><th></th></tr></thead><tbody>${linhas}</tbody></table></div>`
        : `<div class="cm-fraco">Ainda não há medições guardadas neste computador. Os arquivos da pasta "medicoes" do programa antigo entram pelo Importar CSV.</div>
           <div class="cm-botoes"><button type="button" class="btn ghost" onclick="Colorimetro.fecharHist()">Fechar</button></div>`}
    </div>`;
  }

  function desenhar() {
    const p = painel();
    if (!p || !visivel()) return;
    restaurarOS();
    sincronizarRibanas();
    const ref = S.padrao;
    const amo = S.atual && S.atual !== ref ? S.atual : null;
    const r = ref && amo ? comparar(ref, amo) : null;
    const nomeModo = l => (l && l.modo != null ? (MODOS[l.modo] || String(l.modo)) : null);
    const modosDif = r && nomeModo(ref) && nomeModo(amo) && ref.modo !== amo.modo;
    const con = !!S.dev, livre = !S.ocupado && !S.conectando;
    const fo = cfg.formula;
    const barra = r
      ? `<div class="cm-de ${r.aprovado ? 'cm-de-ok' : 'cm-de-fora'}">${fo}=${num(r.des[fo])} ${r.aprovado ? 'Aprovado' : 'Reprovado'}</div>
         <div class="cm-fraco cm-centro">${r.desc.length ? 'Amostra ' + r.desc.join(', ') : 'Dentro da tolerância em L*, a* e b*'}</div>`
      : `<div class="cm-de">${fo}=**</div><div class="cm-fraco cm-centro">${!con ? 'Conecte o aparelho para medir' : (S.os && S.av ? 'Meça a referência e depois a amostra' : 'Meça o padrão e depois a amostra')}</div>`;
    const info = con && S.dev.info
      ? `Modelo: ${escH(S.dev.info.modelo || '?')} · Série: ${escH(S.dev.info.serie)}<br>Firmware: ${escH(S.dev.info.firmware)} · Código ${S.dev.info.codigo} · Modo ${MODOS[modoAtual()] || modoAtual()}`
      : '—';
    const fmtCal = (ok, ts) => (ok ? '✔ ' : '✘ ') + (ts ? new Date(ts * 1000).toLocaleString('pt-BR').slice(0, 17) : 'nunca');
    const cal = S.cal ? `Branco: ${fmtCal(S.cal.branco_ok, S.cal.branco_data)}<br>Preto: ${fmtCal(S.cal.preto_ok, S.cal.preto_data)}` : '—';
    const motivo = semBluetooth();
    const naOS = !!(S.os && S.av);
    /* UM QUADRO SÓ (20/10/2026, Junior: "todos os botões de testar cor, medir
       referência e mudar de amostra devem estar no mesmo quadro. Retire o
       gráfico de comparação de cores"). Na avaliação de uma OS, conectar,
       calibrar e o recado da última medição entram no fim do quadro dela, numa
       linha só; os "Medir padrão / Medir amostra" soltos saem, porque lá quem
       mede é o "Medir amostra N" e o "Amostra N como referência". O gráfico de
       refletância espectral saiu de vez. */
    const aparelho = `
        <div class="cm-botoes ${naOS ? 'cm-aparelho-linha' : ''}">
          ${con
            ? `<button type="button" class="btn" onclick="Colorimetro.desconectar()" ${livre ? '' : 'disabled'}>Desconectar</button>`
            : `<button type="button" class="btn primary" onclick="Colorimetro.conectar(false)" ${livre && !motivo ? '' : 'disabled'}>⇄ Conectar aparelho</button>
               <button type="button" class="btn ghost" onclick="Colorimetro.conectar(true)" ${livre ? '' : 'disabled'} title="Um DS-01 de mentira, com cores aleatórias, para conhecer a tela">Testar sem aparelho</button>`}
          ${naOS ? '' : `<button type="button" class="btn accent" onclick="Colorimetro.medir('padrao')" ${con && livre ? '' : 'disabled'}>Medir padrão (F6)</button>
          <button type="button" class="btn accent" onclick="Colorimetro.medir('amostra')" ${con && livre ? '' : 'disabled'}>Medir amostra (F5)</button>`}
          <button type="button" class="btn small" onclick="Colorimetro.calibrar('branco')" ${con && livre ? '' : 'disabled'} title="Com o aparelho sobre a placa de calibração BRANCA">Calibrar branco</button>
          <button type="button" class="btn small" onclick="Colorimetro.calibrar('preto')" ${con && livre ? '' : 'disabled'} title="Com o aparelho sobre a cavidade/placa PRETA (ou apontado para o vazio)">Calibrar preto</button>
          ${S.simulado && con ? `<button type="button" class="btn small ghost" onclick="Colorimetro.simularBotao()">Simular botão do aparelho</button>` : ''}
        </div>
        <div class="cm-status"><b>${escH(S.status)}</b>${S.conectando || S.ocupado ? ' <span class="cm-fraco">…</span>' : ''}</div>
        ${motivo && !con ? `<div class="cm-aviso">${escH(motivo)} O "Testar sem aparelho" funciona assim mesmo.</div>` : ''}
        ${S.aviso ? `<div class="cm-aviso">${escH(S.aviso)}</div>` : ''}
        ${S.semBD ? `<div class="cm-aviso">Este navegador não deixou guardar o histórico. Exporte o CSV antes de fechar a tela.</div>` : ''}`;
    /* A COMPARAÇÃO DENTRO DO QUADRO DA OS (20/10/2026, Junior: "usuário não
       consegue visualizar comparação de cores sem rolar a página para baixo").
       Na avaliação, o resultado (dE, Aprovado/Reprovado), as duas cores e a
       tabela L*a*b* vão numa coluna à direita do quadro da OS, ao lado dos
       botões; o texto de apresentação da página some. Aparelho, calibração e a
       lista da sessão continuam abaixo — são consulta, não o passo a passo. */
    const comparacao = naOS ? htmlVisorOS() : `
          <div class="cm-blocos">${blocoCor(ref, naOS ? 'Referência' : 'Padrão', 'padrao')}${blocoCor(amo, 'Amostra', 'amostra')}</div>
          ${barra}
          <div class="cm-tab-caixa">${tabelaPagina(ref, amo, r)}</div>
          ${modosDif ? `<div class="cm-aviso">Atenção: o padrão foi medido em ${escH(nomeModo(ref))} e a amostra em ${escH(nomeModo(amo))}. Meça os dois no mesmo modo — a comparação não vale.</div>` : ''}
          <div class="cm-fraco cm-centro cm-peq">${escH(nomeModo(amo) || nomeModo(ref) || MODOS[modoAtual()] || 'SCI')}/${escH(cfg.iluminante1)}°</div>`;
    const sec = p.closest('section');
    if (sec) sec.classList.add('cm-modo-os'); // a apresentação da página some: o quadro precisa caber sem rolar
    p.innerHTML = `
      ${naOS ? htmlAvaliacao(aparelho, comparacao)
        : `<div class="card cm-barra"><div class="cm-os-grade"><div class="cm-os-esq">${htmlEscolherOS()}${aparelho}</div><div class="cm-os-dir">${comparacao}</div></div></div>`}
      ${S.pedido && !S.pedido.lugar ? htmlPedido() : ''}
      ${htmlRel()}${htmlCfg()}${htmlHist()}
      <div class="cm-colunas cm-colunas-os">
        <div class="cm-dir">
          <div class="card cm-aparelho">
            <div><div class="card-title">Aparelho</div><div class="cm-peq">${info}</div></div>
            <div><div class="card-title">Calibração</div><div class="cm-peq">${cal}</div></div>
            <div class="cm-peq" style="grid-column:1/-1">${naOS
              ? `Botão do aparelho: <b>mede a amostra ${S.av.n != null ? S.av.n : 'atual'} como referência para todas as amostras da fase</b>`
              : `Botão do aparelho mede como:
              <label><input type="radio" name="cm-botao" value="amostra" ${cfg.botao === 'amostra' ? 'checked' : ''} onchange="Colorimetro.botaoComo(this.value)"> Amostra</label>
              <label><input type="radio" name="cm-botao" value="padrao" ${cfg.botao === 'padrao' ? 'checked' : ''} onchange="Colorimetro.botaoComo(this.value)"> Padrão</label>`}
              ${cfg.media > 1 ? `<span class="cm-fraco"> · média de ${cfg.media} leituras</span>` : ''}
            </div>
          </div>
          <div class="card">
            <div class="cm-botoes cm-entre">
              <div class="card-title" style="margin:0">Medições desta sessão</div>
              <div class="cm-botoes">
                <button type="button" class="btn small ${S.janela === 'hist' ? 'primary' : ''}" onclick="Colorimetro.alternar('hist')">Histórico</button>
                <button type="button" class="btn small" onclick="Colorimetro.exportar()">Exportar CSV</button>
                <label class="btn small" title="Os arquivos medicoes_AAAA-MM-DD.csv do programa antigo">Importar CSV
                  <input type="file" accept=".csv,text/csv" hidden onchange="Colorimetro.importar(this.files[0]); this.value=''"></label>
                <button type="button" class="btn small ${S.janela === 'cfg' ? 'primary' : ''}" onclick="Colorimetro.alternar('cfg')">Configurações</button>
                <button type="button" class="btn small ghost" onclick="Colorimetro.limpar()">Limpar lista</button>
              </div>
            </div>
            ${S.pedido && S.pedido.lugar === 'lista' ? htmlPedido() : ''}
            <div class="cm-rolar"><table class="cm-lista"><thead><tr><th>Nome</th><th>Hora</th><th>L*</th><th>a*</th><th>b*</th><th>C*</th><th>h°</th>
              <th>dE*ab</th><th>dE00</th><th>Resultado</th><th></th></tr></thead>
              <tbody>${S.leituras.length ? S.leituras.map(linhaLista).join('')
                : '<tr><td colspan="11" class="cm-fraco">As medições aparecem aqui e vão para o histórico deste computador na hora.</td></tr>'}</tbody></table></div>
          </div>
          ${S.log.length ? `<details class="cm-log"><summary>Registro da conexão</summary><pre>${escH(S.log.join('\n'))}</pre></details>` : ''}
        </div>
      </div>`;
  }

  const ESTILO = `
    .cm-botoes{display:flex;flex-wrap:wrap;gap:8px;align-items:flex-end}
    .cm-entre{justify-content:space-between;align-items:center;margin-bottom:8px}
    .cm-status{margin-top:8px;font-size:13px}
    .cm-aviso{margin-top:8px;padding:8px 10px;background:var(--alert-bg);color:var(--alert);border-radius:2px;font-size:13px}
    .cm-colunas{display:grid;grid-template-columns:minmax(300px,420px) 1fr;gap:12px;align-items:start;margin-top:12px}
    @media (max-width:900px){.cm-colunas{grid-template-columns:1fr}}
    .cm-esq,.cm-dir{display:flex;flex-direction:column;gap:12px;min-width:0}
    .cm-esq .card,.cm-dir .card{margin:0}
    .cm-abas{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:8px}
    .cm-tit{font-size:12px;color:var(--ink-3);margin-bottom:4px}
    .cm-tab{width:100%;border-collapse:collapse;font-size:16px}
    .cm-tab th{font-weight:500;color:var(--ink-3);text-align:right;padding:4px 6px}
    .cm-tab td{text-align:right;padding:6px;font-family:'IBM Plex Mono',monospace}
    .cm-tab td.cm-rot{text-align:left;font-family:inherit;color:var(--ink-2)}
    .cm-ok{color:var(--ok)!important}.cm-fora{color:var(--alert)!important}.cm-fraco{color:var(--ink-3)}
    .cm-centro{text-align:center}
    .cm-de{margin:8px 0 4px;padding:10px;text-align:center;font-size:20px;font-weight:700;border-radius:2px;background:var(--line-2)}
    .cm-de-ok{background:#40b12b;color:#fff}.cm-de-fora{background:#b6101f;color:#fff}
    .cm-blocos{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px}
    .cm-refs{display:flex;flex-wrap:wrap;gap:4px;max-height:30vh;overflow:auto}
    .cm-ref{display:flex;align-items:center;gap:4px;padding:2px 6px;border:1px solid var(--line);border-radius:3px;font-size:13px;font-weight:600}
    .cm-ref small{font-weight:400;font-size:10px;color:var(--ink-3)}
    .cm-ref-ok{background:var(--ok-bg);color:var(--ok)}
    .cm-ref-fora{background:var(--alert-bg);color:var(--alert)}
    .cm-ref-cor{width:12px;height:12px;border-radius:2px;border:1px solid var(--line);display:inline-block}
    .cm-bloco{height:150px;border-radius:4px;display:flex;align-items:flex-end;justify-content:center;padding:10px;font-size:15px;cursor:pointer;border:1px solid var(--line)}
    .cm-bloco-vazio{background:repeating-linear-gradient(45deg,var(--line-2),var(--line-2) 8px,transparent 8px,transparent 16px);color:var(--ink-3)}
    .cm-aparelho{display:grid;grid-template-columns:1fr 1fr;gap:10px}
    .cm-peq{font-size:12px;line-height:1.6}
    .cm-peq label{margin-left:10px}
    .cm-rolar{max-height:46vh;overflow:auto}
    .cm-tc{margin-top:10px}
    .cm-tc-tab{border-collapse:collapse;font-size:12px}
    .cm-tc-tab td,.cm-tc-tab th{border:1px solid var(--line);padding:2px 5px;text-align:center;min-width:22px}
    .cm-tc-tab thead th{position:sticky;top:0;background:var(--paper);z-index:1}
    .cm-tc-tab .cm-tc-tom{text-align:left;white-space:nowrap;font-weight:600;border:0;padding-right:8px}
    .cm-tc-tab .cm-tc-eu{color:var(--ink-3)}
    .cm-tc-tab .cm-tc-mesmo{background:var(--ok-bg)}
    .cm-tc-tab .cm-tc-novo-col{border-left:2px solid var(--ink-3)}
    .cm-tc-tab tr.cm-tc-novo-tom td,.cm-tc-tab tr.cm-tc-novo-tom th{border-top:2px solid var(--ink-3)}
    .cm-lista{width:100%;border-collapse:collapse;font-size:12px}
    .cm-lista th{position:sticky;top:0;background:var(--paper);text-align:left;font-weight:500;color:var(--ink-3);padding:4px 6px;border-bottom:1px solid var(--line)}
    .cm-lista td{padding:4px 6px;border-bottom:1px solid var(--line-2);white-space:nowrap}
    .cm-lista td.cm-n{text-align:right;font-family:'IBM Plex Mono',monospace}
    .cm-lista tbody tr{cursor:pointer}
    .cm-lin-padrao{background:var(--line-2)}
    .cm-lin-atual td{font-weight:600}
    .cm-acoes{white-space:nowrap}.cm-acoes .btn{padding:2px 6px}
    .cm-grade{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:8px 12px}
    .cm-campo{display:flex;flex-direction:column;gap:2px;font-size:12px;color:var(--ink-3)}
    .cm-campo input,.cm-campo select{font-size:14px}
    .cm-sub{margin:12px 0 6px;font-weight:600;font-size:13px}
    .cm-cfg,.cm-hist,.cm-pedido,.cm-rel{margin-top:12px}
    .cm-os{border-left:3px solid var(--accent);margin-bottom:12px}
    .cm-os-grade{display:grid;grid-template-columns:minmax(0,1fr) minmax(280px,360px);gap:16px;align-items:start}
    .cm-os-esq{min-width:0}
    .cm-os-dir{border-left:1px solid var(--line);padding-left:16px}
    .cm-os-dir .cm-blocos{margin-top:0;gap:8px}
    .cm-av-topo{margin-bottom:8px}
    .cm-av-num-grande{font-size:30px;line-height:1.1}
    .cm-os-dir .cm-bloco{height:72px;font-size:13px;padding:6px}
    .cm-os-dir .cm-de{font-size:17px;padding:6px;margin:8px 0 2px}
    .cm-os-dir .cm-tab-caixa{margin-top:8px}
    .cm-os-dir .cm-abas .btn{padding:2px 6px;font-size:11px}
    .cm-os-dir .cm-tab{font-size:13px}
    .cm-os-dir .cm-tab td,.cm-os-dir .cm-tab th{padding:2px 4px}
    @media (max-width:1100px){.cm-os-grade{grid-template-columns:1fr}.cm-os-dir{border-left:0;padding-left:0;border-top:1px solid var(--line);padding-top:10px}}
    .cm-colunas-os{grid-template-columns:1fr}
    .cm-modo-os .page-header .desc{display:none}
    .cm-modo-os .page-header{margin-bottom:8px;padding-bottom:6px}
    .cm-av{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:8px;margin:8px 0}
    .cm-av-bloco{padding:6px 10px;background:var(--line-2);border-radius:2px}
    .cm-aparelho-linha{margin-top:10px;padding-top:8px;border-top:1px dashed var(--line)}
    .cm-aparelho-linha .btn{padding:4px 10px;font-size:12px}
    .cm-av-atual{outline:2px solid var(--accent)}
    .cm-av-rot{font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3)}
    .cm-av-num{font-size:20px;font-weight:700;margin:0}
    .cm-fases{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
    .cm-fase{padding:4px 10px;border:1px solid var(--line);border-radius:2px;font-size:12px;color:var(--ink-3)}
    .cm-fase-atual{border-color:var(--accent);color:var(--ink);font-weight:600;outline:1px solid var(--accent)}
    .cm-fase-feita{color:var(--ok)}
    .cm-fase-sem{border-style:dashed;opacity:.6}
    .cm-rel-fase{margin-top:10px;padding-top:6px;border-top:1px solid var(--line-2)}
    .cm-amostra-cor{display:inline-block;width:14px;height:14px;border:1px solid var(--line);vertical-align:-2px;margin-right:6px}
    .cm-tom{padding:6px 8px;border-left:3px solid var(--ok);margin:4px 0;background:var(--line-2);font-size:13px}
    .cm-tom-ruim{border-left-color:var(--alert)}
    .cm-matriz{border-collapse:collapse;font-size:11px}
    .cm-matriz th,.cm-matriz td{padding:3px 6px;border:1px solid var(--line-2);text-align:right;white-space:nowrap}
    .cm-matriz thead th{font-weight:500;color:var(--ink-3)}
    .cm-matriz tbody th{text-align:left;font-weight:500}
    .cm-mz-ok{color:var(--ok)}.cm-mz-fora{color:var(--alert)}.cm-mz-eu{color:var(--ink-3);text-align:center}
    .cm-mz-mesmo{font-weight:700;background:var(--line-2)}
    .cm-previa{padding:8px 10px;border:1px dashed var(--line);font-size:13px;margin-bottom:4px;white-space:pre-wrap}
    .cm-janela-topo{display:flex;justify-content:space-between;align-items:flex-start;gap:8px}
    .cm-janela-topo .card-title{flex:1}
    .cm-pedido{border-left:3px solid var(--accent)}
    .cm-pedido-campo{margin-top:8px;width:100%;max-width:320px;font-size:14px}
    .cm-pedido .cm-botoes{margin-top:12px}
    .cm-cfg .cm-botoes,.cm-hist .cm-botoes{margin-top:12px}
    .cm-log{font-size:11px}.cm-log pre{max-height:200px;overflow:auto;white-space:pre-wrap}
  `;

  function instalar() {
    const st = document.createElement('style');
    st.textContent = ESTILO;
    document.head.appendChild(st);
    const sec = document.querySelector('section.page[data-page="colorimetro"]');
    if (!sec) return;
    let estava = !sec.classList.contains('hidden');
    new MutationObserver(() => {
      const agora = !sec.classList.contains('hidden');
      if (agora && !estava) {
        desenhar();
        carregarTabelas().then(() => { S.leituras.forEach(recalcular); desenhar(); })
          .catch(e => { S.aviso = e.message; desenhar(); });
      }
      estava = agora;
    }).observe(sec, { attributes: true, attributeFilter: ['class'] });
    if (estava) { desenhar(); carregarTabelas().then(desenhar).catch(() => {}); }
    // F5 amostra / F6 padrão, como no original — só nesta tela e com o aparelho
    // conectado (fora dela, e com Ctrl, o F5 continua recarregando a página).
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape' && visivel() && (S.janela || S.pedido)) { S.janela = null; S.pedido = null; desenhar(); return; }
      if (!visivel() || !S.dev || e.ctrlKey || e.altKey || e.metaKey || e.shiftKey) return;
      if (e.key === 'F5' || e.key === 'F6') {
        e.preventDefault();
        medir(e.key === 'F6' ? 'padrao' : 'amostra');
      }
    });
  }

  window.Colorimetro = {
    conectar, desconectar, calibrar, medir, selecionar, usarPadrao, renomear, excluir, limpar, exportar, importar,
    abrirHistorico: () => alternar('hist', true), carregarDia: d => carregarDia(d).catch(e => { S.aviso = e.message; desenhar(); }),
    histUsarPadrao, histExportar, fecharHist: fecharJanela, fecharCfg: fecharJanela, fecharJanela, alternar,
    abrirCfg: () => alternar('cfg', true),
    aplicarCfg, resetarCfg, responder, avaliarOS, encerrarOS, relMudar, relSalvar, exportarAvaliacao, avTrocar, avIrPara, avDefinirTotal, avConcluirFase, avReavaliarFase, avLimpar,
    pagina: i => { S.pagina = i; desenhar(); },
    botaoComo: v => { cfg.botao = v === 'padrao' ? 'padrao' : 'amostra'; salvarCfg(); },
    // Como o DS-01 de verdade: o botão mede no modo do aparelho (SCI), não no escolhido.
    simularBotao: () => { if (S.t && S.t.apertarBotao) S.t.apertarBotao(0); },
    // para os testes: as contas, sem tela
    _contas: { carregarTabelas, espectroParaLab, deAb, de2000, de94, deCmc, labHex, comparar, arred, num, pacote, checksumOk }
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', instalar);
  else instalar();
})();
