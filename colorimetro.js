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
    iluminante1: 'D65/10', iluminante2: 'A/2', modo: 'auto', formula: 'dE*ab', media: 1,
    cmc_l: 2, cmc_c: 1,
    tol: { dL: 0.4, da: 0.1, db: 0.5, 'dE*ab': 0.6, dC: 0.6, dH: 0.6, "dL'": 0.6, "dC'": 0.6, "dH'": 0.6,
      dE00: 0.6, dEcmc: 0.6, dE94: 0.6 },
    nome_padrao: 'Padrão', nome_amostra: 'Amostra', k94: [1, 1, 1], k00: [1, 1, 1],
    botao: 'amostra', ultimo_aparelho: null
  };

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
    // Como o app: aprova/reprova só pelo dE escolhido, com o valor exato.
    const aprovado = des[cfg.formula] <= tol[cfg.formula];
    // As dicas por coluna saem dos valores já arredondados na tela; não entram na aprovação.
    const d1 = [0, 1, 2].map(i => arred(l2[i]) - arred(l1[i]));
    const rot = [['Claro', 'Escuro'], ['Vermelho', 'Verde'], ['Amarelo', 'Azul']];
    const txt = [['mais clara', 'mais escura'], ['mais vermelha', 'mais verde'], ['mais amarela', 'mais azul']];
    const dicas = [], desc = [];
    ['dL', 'da', 'db'].forEach((k, i) => {
      const v = d1[i];
      if (dentro(v, tol[k])) dicas.push(['Aprovado', true]);
      else { dicas.push([v > 0 ? rot[i][0] : rot[i][1], false]); desc.push(v > 0 ? txt[i][0] : txt[i][1]); }
    });
    return { d: [dL, da, db, dC, dH], d1, dicas, des, aprovado, desc };
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

  function calibrar(tipo) {
    if (!S.dev || S.ocupado) return;
    perguntar({
      titulo: `Calibrar ${tipo}`,
      texto: tipo === 'branco'
        ? 'Coloque o aparelho sobre a placa de calibração BRANCA e clique em Calibrar.'
        : 'Coloque o aparelho sobre a cavidade/placa de calibração PRETA (ou aponte para o vazio) e clique em Calibrar.',
      sim: 'Calibrar', executar: () => calibrarAgora(tipo)
    });
  }
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
    if (S.os && S.av && S.av.n == null) { S.aviso = 'A rodada terminou: clique em Reavaliar ou escolha o número da amostra.'; desenhar(); return; }
    S.ocupado = true;
    status('Medindo…');
    try { chegou(await S.dev.medir(modoAtual()), tipo); } catch (e) { erroDoAparelho(e); }
  }

  // O botão do próprio DS-01 foi apertado: a medição já foi feita, só baixar.
  function aoBotao(modo) {
    if (S.ocupado || !S.dev) return;
    let tipo = cfg.botao;
    if (tipo === 'amostra' && !S.padrao && !S.acumulando.length) tipo = 'padrao';
    // Na avaliação de uma OS o botão mede a amostra atual (a 1ª da rodada vira a referência).
    if (S.os && S.av) { if (S.av.n == null) return; tipo = 'amostra'; }
    S.ocupado = true;
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
      lab_aparelho: labAp, horario: new Date() });
    // Com a tela ligada a uma OS, quem dá nome, referência e tom é a avaliação.
    const msg = naOS ? avMedicao(l, tipo) : '';
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
      refId: l.refId, refNome: ref ? ref.nome : '', iluminante: cfg.iluminante1,
      osId: l.osId || null, osNumero: l.osNumero || '', amostraN: l.amostraN || null, rodada: l.rodada || null };
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
        S.padrao = S.os && S.av ? refComoLeitura(S.av.rodada) : null;
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
      lab_aparelho: r.lab_aparelho, horario: new Date(r.horario) });
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
      lab_aparelho: r.lab_aparelho, horario: r.horario });
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
  function avNovo(osId) { return { osId, n: 1, rodada: 1, tons: {}, ref: {}, medidas: {}, fila: null, fimRodada: false }; }
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
    try { localStorage.setItem(CHAVE_AV(S.avOS.osId), JSON.stringify(S.avOS)); } catch (e) { /* sem storage */ }
  }
  // As fases de enfesto da OS, com o tecido e a cor de cada uma.
  function fasesDaOS(osId) {
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
      return { i, nome: f.nome || `Fase ${i + 1}`, tecido, rotulo: `Fase ${i + 1} · ${f.nome || 'sem nome'}${tecido ? ' · ' + tecido : ''}` };
    });
  }
  // Põe a fase i na tela: S.av passa a ser a avaliação dela.
  function avUsarFase(i) {
    const c = S.avOS;
    c.faseAtual = i;
    if (!c.fases[i]) c.fases[i] = avNovo(c.osId);
    S.av = c.fases[i];
    S.padrao = refComoLeitura(S.av.rodada);
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

  // A referência do tom como "leitura", para a comparação da tela (barra, blocos, lista).
  function refComoLeitura(k) {
    const r = S.av && S.av.ref[k];
    if (!r) return null;
    const ja = S.leituras.find(x => x.id === r.medId);
    return ja || { id: r.medId, nome: r.nome, tipo: 'padrao', lab1: r.lab, lab2: r.lab2 || null, espectro: [], horario: new Date(r.em) };
  }

  function avaliarOS(id) {
    const o = (typeof STATE !== 'undefined' && STATE && STATE.ordens || []).find(x => x.id === id);
    if (!o) return;
    S.os = { id: o.id, numero: o.os || '', rotulo: rotuloOS(o) };
    S.avOS = avCarregar(o.id);
    S.fases = fasesDaOS(o.id);
    avUsarFase(Math.min(S.avOS.faseAtual || 0, S.fases.length - 1));
    S.janela = null;
    S.aviso = '';
    if (typeof goto === 'function') goto('colorimetro');
    status(S.dev ? avDica() : `OS ${S.os.numero} escolhida — conecte o aparelho para medir`);
  }
  function encerrarOS() {
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
    if (av.fimRodada) return `Fim da rodada do Tom ${av.rodada}. ${avPendentes().length ? 'Clique em Reavaliar para formar o tom seguinte.' : 'Nenhuma amostra a reavaliar.'}`;
    if (!av.ref[av.rodada]) return `Tom ${av.rodada}: meça a Amostra ${av.n} (ela vira a referência do tom) ou vá ao número da amostra que vai ser a referência.`;
    return `Tom ${av.rodada}: meça a Amostra ${av.n}; uma aprovação basta. Clique em Trocar para passar à seguinte.`;
  }
  const avPendentes = (av = S.av) => numerosDasAmostras(av).filter(n => tomDe(n, av) === 'r');

  // Uma medição chegou com a tela ligada a uma OS (chamado por chegou()).
  function avMedicao(l, tipo) {
    const av = S.av, k = av.rodada;
    l.osId = S.os.id; l.osNumero = S.os.numero; l.rodada = k;
    l.fase = S.avOS.faseAtual; l.faseNome = (S.fases[S.avOS.faseAtual] || {}).rotulo || '';
    if (av.fimRodada || av.n == null) {
      return 'A rodada terminou: clique em Reavaliar (ou escolha o número da amostra) antes de medir.';
    }
    if (tipo === 'padrao') return avReferencia(l);
    const n = av.n;
    l.amostraN = n;
    l.nome = `${cfg.nome_amostra} ${n}`;
    const lista = av.medidas[n] = av.medidas[n] || [];
    if (!av.ref[k]) {
      // A primeira medição da rodada define o tom: é a referência, e a amostra já está nele.
      av.ref[k] = { medId: l.id, amostra: n, nome: l.nome, lab: l.lab1, lab2: l.lab2, em: l.horario.toISOString() };
      av.tons[n] = k;
      lista.push({ id: l.id, lab: l.lab1, rodada: k, de: 0, aprovado: true, ref: true, em: l.horario.toISOString() });
      S.padrao = l;
      avGuardar();
      return `Amostra ${n} é a referência do Tom ${k}. Clique em Trocar para passar à Amostra ${proximoNumero(n) || n + 1}.`;
    }
    const ref = refComoLeitura(k);
    const r = comparar(ref, l);
    l.refId = ref.id;
    const de = r.des[cfg.formula];
    lista.push({ id: l.id, lab: l.lab1, rodada: k, de, aprovado: r.aprovado, em: l.horario.toISOString() });
    let msg;
    if (r.aprovado && !estaNumTom(n)) {
      av.tons[n] = k;
      msg = `✔ Amostra ${n} aprovada (${cfg.formula} ${num(de)}) → entrou no Tom ${k}.`;
    } else if (r.aprovado) {
      msg = `✔ Amostra ${n} aprovada de novo (${cfg.formula} ${num(de)}); continua no Tom ${tomDe(n)}.`;
    } else if (estaNumTom(n)) {
      msg = `Amostra ${n} reprovada nesta leitura (${cfg.formula} ${num(de)}), mas já tem aprovação: continua no Tom ${tomDe(n)}.`;
    } else {
      msg = `✘ Amostra ${n} reprovada (${cfg.formula} ${num(de)}). Meça de novo ou clique em Trocar (ela vai para a reavaliação).`;
    }
    avGuardar();
    return msg;
  }

  /* A REFERÊNCIA É UMA AMOSTRA (20/10/2026, Junior: "as amostras são usadas
     como referência. Logo, se a amostra n1 é utilizada como referência as
     demais amostras devem ser definidas como tom 1. Se selecionar amostra n5
     como referência, as amostras compatíveis devem ser definidas como tom 2,
     pois houve repescagem das amostras que não foram compatíveis com Tom 1").
     "Medir referência" media um padrão sem número ("Referência do Tom 2"),
     solto das amostras. Agora mede a AMOSTRA ATUAL e ela vira a referência:
       · o tom em formação ainda sem referência de amostra → é a dele (Amostra 1
         na 1ª rodada = Tom 1; na repescagem, a escolhida = Tom 2…);
       · o tom já tem a referência em OUTRA amostra → é repescagem: abre o tom
         seguinte com ela, e a fila são as que não entraram em tom nenhum;
       · a mesma amostra de novo → só troca a leitura da referência. */
  function avReferencia(l) {
    const av = S.av, n = av.n;
    const atual = av.ref[av.rodada];
    const repescagem = !!(atual && atual.amostra != null && atual.amostra !== n);
    if (repescagem) {
      const fora = numerosDasAmostras(av).filter(x => x !== n && !estaNumTom(x) && tomDe(x) !== 0);
      av.rodada = tomNovoDe(av);
      av.fila = [n, ...fora].sort((a, b) => a - b);
      av.fimRodada = false;
    }
    const k = av.rodada;
    l.tipo = 'amostra'; l.amostraN = n; l.rodada = k;
    l.nome = `${cfg.nome_amostra} ${n}`;
    const em = l.horario.toISOString();
    av.ref[k] = { medId: l.id, amostra: n, nome: l.nome, lab: l.lab1, lab2: l.lab2, em };
    av.tons[n] = k;
    (av.medidas[n] = av.medidas[n] || []).push({ id: l.id, lab: l.lab1, rodada: k, de: 0, aprovado: true, ref: true, em });
    S.padrao = l;
    avGuardar();
    const prox = proximoNumero(n);
    return `Amostra ${n} é a referência do Tom ${k}${repescagem ? ` (repescagem: ${av.fila.filter(x => x !== n).join(', ') || 'nenhuma outra'})` : ''}.`
      + (prox != null ? ` Clique em Trocar para passar à Amostra ${prox}.` : '');
  }

  // O número seguinte: na 1ª rodada, o próximo inteiro (amostras novas); nas
  // seguintes, a próxima da fila de reavaliação ainda não medida nesta rodada —
  // dando a volta, porque a referência escolhida pode estar no meio da fila
  // (Amostra 5 com 3 e 4 antes dela).
  function proximoNumero(n) {
    const av = S.av;
    if (av.rodada === 1 && !av.fila) return n + 1;
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
    const medidasDaRodada = (av.medidas[n] || []).filter(m => m.rodada === av.rodada);
    let msg = '';
    if (!estaNumTom(n) && medidasDaRodada.length) { av.tons[n] = 'r'; msg = `Amostra ${n} foi para a reavaliação. `; }
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
    S.av.n = n; S.av.fimRodada = false;
    S.atual = null;
    avGuardar();
    status(`Agora é a Amostra ${n}${estaNumTom(n) ? ` (já está no Tom ${tomDe(n)})` : tomDe(n) === 'r' ? ' (está a reavaliar)' : ''}.`);
  }
  function avReavaliar() {
    const av = S.av;
    const fila = avPendentes();
    if (!av || !fila.length) return;
    av.rodada = Math.max(av.rodada, ...Object.keys(av.ref).map(Number)) + 1;
    av.fila = fila;
    av.n = fila[0];
    av.fimRodada = false;
    S.padrao = null; S.atual = null;
    avGuardar();
    status(`Rodada do Tom ${av.rodada}: ${fila.length} amostras a reavaliar (${fila.join(', ')}). Meça a Amostra ${av.n} — ela vira a referência do Tom ${av.rodada} — ou vá ao número que vai ser a referência.`);
  }
  // Concluir a fase atual e passar para a seguinte (as fases andam em ordem).
  function avConcluirFase() {
    const c = S.avOS;
    if (!c) return;
    const i = c.faseAtual;
    if (!numerosDasAmostras().length) { S.aviso = 'Meça ao menos uma amostra desta fase antes de concluí-la.'; desenhar(); return; }
    c.concluidas[i] = true;
    const prox = S.fases.findIndex((f, j) => j > i && !c.concluidas[j]);
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
        status(`Reavaliação da ${S.fases[i].rotulo}: meça a Amostra 1 (ela vira a referência do Tom 1).`);
      } });
  }

  /* ---------------- o relatório ---------------- */
  // Os tons em uso, na ordem dos números das rodadas.
  function tonsDoRelatorio(av = S.av) {
    const nums = numerosDasAmostras(av);
    const ks = [...new Set(nums.map(n => tomDe(n, av)).filter(t => typeof t === 'number' && t > 0))].sort((a, b) => a - b);
    return ks.map(k => {
      const membros = nums.filter(n => tomDe(n, av) === k);
      const ref = av.ref[k];
      // Cada amostra pela sua melhor medição contra a referência deste tom.
      const melhor = n => {
        const ms = (av.medidas[n] || []).filter(m => m.rodada === k);
        return ms.length ? ms.reduce((a, b) => (a.de <= b.de ? a : b)) : (av.medidas[n] || []).slice(-1)[0];
      };
      const ruins = [];
      let pior = 0;
      membros.forEach(n => {
        const m = melhor(n);
        if (!m || !ref) return;
        const d = deDaFormula(ref.lab, m.lab);
        pior = Math.max(pior, d);
        if (d > tolAtual() + 1e-9) ruins.push([n, d]);
      });
      return { k, membros, ref, pior, ruins, melhor };
    });
  }
  const nomeAm = n => `${cfg.nome_amostra} ${n}`;
  // As fases que já têm medição, na ordem da OS.
  const fasesComMedicao = () => (S.fases || []).filter(f => S.avOS.fases[f.i] && numerosDasAmostras(S.avOS.fases[f.i]).length);
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
       Corpo · Malha Algodão Vermelho (1,2,3,4,5,6) - Gola · Ribana (1) */
  function textoFaseCurto(f) {
    const av = S.avOS.fases[f.i];
    const grupos = tonsDoRelatorio(av).map(t => `(${t.membros.join(',')})`).join(' ');
    const pend = avPendentes(av), fora = numerosDasAmostras(av).filter(n => tomDe(n, av) === 0);
    const nome = [f.nome, f.tecido].filter(Boolean).join(' · ');
    return `${nome}${S.avOS.reav[f.i] ? ' [reavaliada]' : ''} ${grupos || '(nenhuma)'}`
      + (pend.length ? `, a reavaliar ${pend.join(',')}` : '')
      + (fora.length ? `, não usar ${fora.join(',')}` : '');
  }
  // Um recado só, com todas as fases medidas, na ordem do enfesto.
  function textoRelatorio() {
    return 'Tecidos compatíveis em cor\n' + fasesComMedicao().map(textoFaseCurto).join(' - ');
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
    if (v === 'novo') {
      const k = tomNovoDe(av);
      const m = (av.medidas[n] || []).slice(-1)[0];
      if (m) av.ref[k] = { medId: m.id, amostra: n, nome: nomeAm(n), lab: m.lab, em: m.em };
      av.tons[n] = k;
      avGuardar();
      desenhar();
      return;
    }
    av.tons[n] = v === 'r' ? 'r' : (parseInt(v, 10) || 0);
    avGuardar();
    desenhar();
  }
  async function relSalvar() {
    if (!S.avOS || !S.os) return;
    if (!fasesComMedicao().some(f => tonsDoRelatorio(S.avOS.fases[f.i]).length)) { S.aviso = 'Nenhuma amostra está num tom ainda.'; desenhar(); return; }
    if (typeof salvarObsNota !== 'function') { S.aviso = 'Não achei a gravação das observações da OS.'; desenhar(); return; }
    const o = (STATE.ordens || []).find(x => x.id === S.os.id);
    if (!o) { S.aviso = `A OS ${S.os.numero} não existe mais.`; desenhar(); return; }
    const notas = () => (Array.isArray(o.obsNotas) ? o.obsNotas.length : 0);
    const antes = notas();
    try { await salvarObsNota(o.id, '', textoRelatorio()); } catch (e) { S.aviso = 'Não deu para salvar: ' + e.message; desenhar(); return; }
    if (notas() <= antes) { S.aviso = 'O programa não gravou a observação (veja o aviso no canto da tela).'; desenhar(); return; }
    S.relSalvo = { osId: o.id, em: new Date() };
    S.aviso = '';
    status(`Relatório salvo nas observações da OS ${o.os}`);
  }

  // As fases em fila, como chips: a atual destacada, as concluídas com ✔.
  function htmlFases() {
    if (!S.fases || S.fases.length < 2) return '';
    const c = S.avOS;
    const chips = S.fases.map(f => {
      const atual = f.i === c.faseAtual, feita = !!c.concluidas[f.i];
      const av = c.fases[f.i];
      const tons = av ? tonsDoRelatorio(av).length : 0;
      return `<span class="cm-fase ${atual ? 'cm-fase-atual' : ''} ${feita ? 'cm-fase-feita' : ''}">${feita ? '✔ ' : ''}${escH(f.rotulo)}`
        + `${av && numerosDasAmostras(av).length ? ` <span class="cm-fraco">· ${tons} ${tons === 1 ? 'tom' : 'tons'}</span>` : ''}${c.reav[f.i] ? ' <span class="cm-fraco">· reavaliada</span>' : ''}</span>`;
    }).join('');
    const feitas = S.fases.filter(f => c.concluidas[f.i]);
    return `<div class="cm-fases">${chips}</div>
      ${feitas.length ? `<div class="cm-botoes" style="margin-top:6px">
        <label class="cm-campo" style="max-width:340px"><span>Reavaliar a fase</span>
          <select onchange="if(this.value!=='')Colorimetro.avReavaliarFase(this.value); this.value=''">
            <option value="">— escolha a fase que vai ser reavaliada —</option>
            ${feitas.map(f => `<option value="${f.i}">${escH(f.rotulo)}</option>`).join('')}</select></label></div>` : ''}`;
  }

  function htmlAvaliacao() {
    if (!S.os || !S.av) return '';
    const av = S.av, k = av.rodada, ref = av.ref[k];
    const con = !!S.dev, livre = !S.ocupado && !S.conectando;
    const concluida = !!S.avOS.concluidas[S.avOS.faseAtual];
    const fase = S.fases[S.avOS.faseAtual];
    const ultimaFase = !S.fases.some((f, j) => j > S.avOS.faseAtual && !S.avOS.concluidas[j]);
    const ms = av.n != null ? (av.medidas[av.n] || []).filter(m => m.rodada === k) : [];
    const ult = ms[ms.length - 1];
    const situacao = av.n == null ? '' : estaNumTom(av.n) ? `<span class="cm-ok">no Tom ${tomDe(av.n)}</span>`
      : tomDe(av.n) === 'r' ? '<span class="cm-fora">a reavaliar</span>' : '<span class="cm-fraco">sem aprovação ainda</span>';
    const pend = avPendentes();
    const tons = tonsDoRelatorio();
    const prox = av.n != null ? proximoNumero(av.n) : null;
    const pode = con && livre && !concluida;
    return `<div class="card cm-os">
      <div class="cm-botoes cm-entre">
        <div><span class="cm-fraco">Avaliando as amostras da</span> <b>${escH(S.os.rotulo)}</b></div>
        <div class="cm-botoes">
          <button type="button" class="btn small ${S.janela === 'rel' ? 'primary' : ''}" onclick="Colorimetro.alternar('rel')">Relatório de tonalidades</button>
          <button type="button" class="btn small ghost" onclick="Colorimetro.encerrarOS()">Medir sem OS</button></div></div>
      ${htmlFases()}
      ${concluida ? `<div class="cm-aviso" style="background:var(--ok-bg);color:var(--ok)">${escH(fase.rotulo)} concluída. Para repetir esta fase, escolha-a em "Reavaliar a fase".</div>` : `
      <div class="cm-av">
        <div class="cm-av-bloco"><div class="cm-av-rot">${S.fases.length > 1 ? escH(fase.rotulo) + ' · ' : ''}Tom em formação</div><div class="cm-av-num">Tom ${k}</div>
          <div class="cm-peq">${ref ? `Referência: <b>${escH(ref.amostra ? nomeAm(ref.amostra) : ref.nome)}</b> · L* ${num(ref.lab[0])} a* ${num(ref.lab[1])} b* ${num(ref.lab[2])}`
            : 'Referência: a primeira medição da rodada'}</div></div>
        <div class="cm-av-bloco cm-av-atual"><div class="cm-av-rot">Amostra atual</div>
          <div class="cm-av-num">${av.n != null ? 'Nº ' + av.n : '—'}</div>
          <div class="cm-peq">${av.n != null ? `${ms.length} ${ms.length === 1 ? 'medição' : 'medições'} nesta rodada${ult && !ult.ref ? ` · última ${escH(cfg.formula)} ${num(ult.de)} ${ult.aprovado ? '✔' : '✘'}` : ''} · ${situacao}` : escH(avDicaFase())}</div></div>
        <div class="cm-av-bloco"><div class="cm-av-rot">Até agora nesta fase</div>
          <div class="cm-peq">${tons.length ? tons.map((t, i) => `<b>Tom ${i + 1}</b>: ${t.membros.join(', ')}`).join('<br>') : 'nenhum tom ainda'}
          ${pend.length ? `<br><span class="cm-fora">A reavaliar: ${pend.join(', ')}</span>` : ''}</div></div>
      </div>
      <div class="cm-botoes">
        <button type="button" class="btn accent" onclick="Colorimetro.medir('amostra')" ${pode && av.n != null ? '' : 'disabled'}>Medir amostra ${av.n != null ? av.n : ''} (F5)</button>
        <button type="button" class="btn primary" onclick="Colorimetro.avTrocar()" ${av.n != null ? '' : 'disabled'}
          title="Passa para a amostra seguinte. Sem aprovação nesta rodada, esta vai para a reavaliação.">Trocar → ${prox != null ? 'amostra ' + prox : 'fim da rodada'}</button>
        <label class="cm-campo" style="max-width:120px"><span>Ir para a amostra nº</span>
          <input type="number" min="1" value="${av.n != null ? av.n : ''}" onchange="Colorimetro.avIrPara(this.value)"></label>
        <button type="button" class="btn small" onclick="Colorimetro.medir('padrao')" ${pode && av.n != null ? '' : 'disabled'}
          title="A amostra atual vira a referência: as compatíveis com ela entram no tom dela. Se o tom em formação já tem a referência em outra amostra, abre o tom seguinte (repescagem das que não entraram em tom nenhum).">Amostra ${av.n != null ? av.n : ''} como referência do Tom ${ref && ref.amostra != null && ref.amostra !== av.n ? tomNovoDe(av) : k} (F6)</button>
        ${pend.length ? `<button type="button" class="btn ${av.fimRodada ? 'primary' : ''}" onclick="Colorimetro.avReavaliar()">Reavaliar ${pend.length} → Tom ${Math.max(k, ...Object.keys(av.ref).map(Number)) + 1}</button>` : ''}
        <button type="button" class="btn ${av.fimRodada && !pend.length ? 'primary' : ''}" onclick="Colorimetro.avConcluirFase()" ${numerosDasAmostras().length ? '' : 'disabled'}
          title="Fecha esta fase e passa para a seguinte do enfesto">${ultimaFase ? 'Concluir a fase' : 'Concluir a fase → próxima'}</button>
      </div>`}</div>`;
  }

  // O relatório de uma fase: amostras, tons e o seletor para corrigir.
  function htmlRelFase(f) {
    const av = S.avOS.fases[f.i], fo = cfg.formula;
    const nums = numerosDasAmostras(av);
    const tons = tonsDoRelatorio(av);
    const ordem = {};
    tons.forEach((t, i) => { ordem[t.k] = i + 1; });
    const ks = Object.keys(av.ref).map(Number).sort((a, b) => a - b);
    const linhas = nums.map(n => {
      const t = tomDe(n, av);
      const todas = av.medidas[n] || [];
      const apr = todas.filter(m => m.aprovado).length;
      const opc = ks.map(k => `<option value="${k}" ${t === k ? 'selected' : ''}>Tom ${ordem[k] || '?'}${ordem[k] ? '' : ' (vazio)'}</option>`)
        .concat([`<option value="novo">Tom ${tons.length + 1} (novo, só esta amostra)</option>`,
          `<option value="r" ${t === 'r' ? 'selected' : ''}>A reavaliar</option>`, `<option value="0" ${t === 0 ? 'selected' : ''}>Não usar</option>`,
          t == null ? '<option value="" selected>—</option>' : '']);
      const ultima = todas[todas.length - 1];
      const semRef = todas.filter(m => !m.ref);
      return `<tr><td>${ultima ? `<span class="cm-amostra-cor" style="background:${labHex(...ultima.lab)}"></span>` : ''}${escH(nomeAm(n))}</td>
        <td class="cm-n">${todas.length}</td><td class="cm-n">${apr}</td>
        <td class="cm-n">${semRef.length ? num(Math.min(...semRef.map(m => m.de))) : (todas.some(m => m.ref) ? 'ref.' : '')}</td>
        <td><select onchange="Colorimetro.relMudar(${f.i}, ${n}, this.value)">${opc.join('')}</select></td></tr>`;
    }).join('');
    const resumo = tons.map((t, i) => `<div class="cm-tom ${t.ruins.length ? 'cm-tom-ruim' : ''}">
        <b>Tom ${i + 1}</b> — ${t.membros.map(nomeAm).join(', ')}
        <span class="cm-fraco">· referência ${escH(t.ref ? (t.ref.amostra ? nomeAm(t.ref.amostra) : t.ref.nome) : '—')}${t.membros.length > 1 ? ` · ${escH(fo)} máx ${num(t.pior)}` : ''}</span>
        ${t.ruins.length ? `<div class="cm-fora">Fora da tolerância da referência: ${t.ruins.map(([n, d]) => `${escH(nomeAm(n))} (${num(d)})`).join('; ')}</div>` : ''}
      </div>`).join('');
    const pend = avPendentes(av);
    const marcas = marcasDaFase(f);
    return `<div class="cm-rel-fase">
      <div class="cm-sub">${escH(S.fases.length > 1 ? f.rotulo : (f.tecido || 'Amostras'))}${marcas.length ? ` <span class="cm-fraco">(${escH(marcas.join(', '))})</span>` : ''}</div>
      <div class="cm-rolar"><table class="cm-lista"><thead><tr><th>Amostra</th><th>Medições</th><th>Aprovações</th><th>Melhor ${escH(fo)}</th><th>Tom</th></tr></thead>
        <tbody>${linhas}</tbody></table></div>
      ${resumo || '<div class="cm-fraco">Nenhuma amostra num tom.</div>'}
      ${pend.length ? `<div class="cm-fora" style="margin-top:6px">A reavaliar: ${pend.map(nomeAm).join(', ')}</div>` : ''}
    </div>`;
  }

  function htmlRel() {
    if (S.janela !== 'rel') return '';
    const topo = `<div class="cm-janela-topo"><div class="card-title">Relatório de tonalidades · ${escH(S.os ? S.os.rotulo : '')}</div>${botaoFechar()}</div>`;
    if (!S.avOS) return `<div class="card cm-rel">${topo}<div class="cm-fraco">Escolha a OS pelo "⋯ → avaliar amostras" da lista de OS.</div></div>`;
    const fo = cfg.formula, tol = tolAtual();
    const fases = fasesComMedicao();
    if (!fases.length) return `<div class="card cm-rel">${topo}<div class="cm-fraco">Ainda não há amostras medidas para a OS ${escH(S.os.numero)}.</div></div>`;
    const faltam = S.fases.filter(f => !S.avOS.concluidas[f.i]);
    const algumTom = fases.some(f => tonsDoRelatorio(S.avOS.fases[f.i]).length);
    const salvo = S.relSalvo && S.relSalvo.osId === S.os.id
      ? `<span class="cm-ok">✔ salvo nas observações às ${horaTxt(S.relSalvo.em).slice(0, 5)}</span>
         <button type="button" class="btn small ghost" onclick="verOS('${escH(S.os.id)}')">Ver a folha da OS</button>` : '';
    return `<div class="card cm-rel">${topo}
      <div class="cm-fraco" style="margin-bottom:8px">Por fase de enfesto e tecido. Cada amostra entrou no tom da rodada em que teve a primeira aprovação contra a referência
        (<b>${escH(fo)}</b> até <b>${numTol(tol)}</b>, ${escH(cfg.iluminante1)}°). Corrija o tom de qualquer amostra aqui; o texto abaixo acompanha.</div>
      ${fases.map(htmlRelFase).join('')}
      ${faltam.length && S.fases.length > 1 ? `<div class="cm-fora" style="margin-top:8px">Fases ainda não concluídas: ${faltam.map(f => escH(f.rotulo)).join('; ')}</div>` : ''}
      <div class="cm-sub">Vai para as observações da OS ${escH(S.os.numero)}</div>
      <div class="cm-previa">${escH(algumTom ? textoRelatorio() : '—')}</div>
      <div class="cm-botoes">
        <button type="button" class="btn primary" onclick="Colorimetro.relSalvar()" ${algumTom ? '' : 'disabled'}>Salvar nas observações da OS</button>
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
    let titulo, cab, linhas;
    if (S.pagina === 0) {
      titulo = `L*a*b* · ${il1}°`; cab = ['L*', 'a*', 'b*'];
      const d = r ? r.d1 : vazio3;
      linhas = [['Padrão', (ref ? ref.lab1 : vazio3).map(v => f(v))], ['Amostra', (amo ? amo.lab1 : vazio3).map(v => f(v))],
        ['Diferença', d.map(v => f(v, true))]];
      if (r) linhas.push(['', r.dicas.map(x => x[0]), r.dicas.map(x => x[1] ? 'cm-ok' : 'cm-fora')]);
    } else if (S.pagina === 1) {
      titulo = `L*C*h° · ${il1}°`; cab = ['L*', 'C*', 'h°'];
      const lc = x => x ? [x.lab1[0], ...lch(...x.lab1)] : vazio3;
      const d = r ? r.d : [null, null, null, null, null];
      linhas = [['Padrão', lc(ref).map(v => f(v))], ['Amostra', lc(amo).map(v => f(v))],
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
      linhas = [['Padrão', (a || vazio3).map(v => f(v))], ['Amostra', (b || vazio3).map(v => f(v))],
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
    const ref = S.padrao;
    const amo = S.atual && S.atual !== ref ? S.atual : null;
    const r = ref && amo ? comparar(ref, amo) : null;
    const con = !!S.dev, livre = !S.ocupado && !S.conectando;
    const fo = cfg.formula;
    const barra = r
      ? `<div class="cm-de ${r.aprovado ? 'cm-de-ok' : 'cm-de-fora'}">${fo}=${num(r.des[fo])} ${r.aprovado ? 'Aprovado' : 'Reprovado'}</div>
         <div class="cm-fraco cm-centro">${r.desc.length ? 'Amostra ' + r.desc.join(', ') : 'Dentro da tolerância em L*, a* e b*'}</div>`
      : `<div class="cm-de">${fo}=**</div><div class="cm-fraco cm-centro">${con ? 'Meça o padrão e depois a amostra' : 'Conecte o aparelho para medir'}</div>`;
    const info = con && S.dev.info
      ? `Modelo: ${escH(S.dev.info.modelo || '?')} · Série: ${escH(S.dev.info.serie)}<br>Firmware: ${escH(S.dev.info.firmware)} · Código ${S.dev.info.codigo} · Modo ${MODOS[modoAtual()] || modoAtual()}`
      : '—';
    const fmtCal = (ok, ts) => (ok ? '✔ ' : '✘ ') + (ts ? new Date(ts * 1000).toLocaleString('pt-BR').slice(0, 17) : 'nunca');
    const cal = S.cal ? `Branco: ${fmtCal(S.cal.branco_ok, S.cal.branco_data)}<br>Preto: ${fmtCal(S.cal.preto_ok, S.cal.preto_data)}` : '—';
    const motivo = semBluetooth();
    p.innerHTML = `
      ${htmlAvaliacao()}
      <div class="card cm-barra">
        <div class="cm-botoes">
          ${con
            ? `<button type="button" class="btn" onclick="Colorimetro.desconectar()" ${livre ? '' : 'disabled'}>Desconectar</button>`
            : `<button type="button" class="btn primary" onclick="Colorimetro.conectar(false)" ${livre && !motivo ? '' : 'disabled'}>⇄ Conectar aparelho</button>
               <button type="button" class="btn ghost" onclick="Colorimetro.conectar(true)" ${livre ? '' : 'disabled'} title="Um DS-01 de mentira, com cores aleatórias, para conhecer a tela">Testar sem aparelho</button>`}
          <button type="button" class="btn accent" onclick="Colorimetro.medir('padrao')" ${con && livre ? '' : 'disabled'}>Medir padrão (F6)</button>
          <button type="button" class="btn accent" onclick="Colorimetro.medir('amostra')" ${con && livre ? '' : 'disabled'}>Medir amostra (F5)</button>
          <button type="button" class="btn small" onclick="Colorimetro.calibrar('branco')" ${con && livre ? '' : 'disabled'}>Calibrar branco</button>
          <button type="button" class="btn small" onclick="Colorimetro.calibrar('preto')" ${con && livre ? '' : 'disabled'}>Calibrar preto</button>
          ${S.simulado && con ? `<button type="button" class="btn small ghost" onclick="Colorimetro.simularBotao()">Simular botão do aparelho</button>` : ''}
        </div>
        <div class="cm-status"><b>${escH(S.status)}</b>${S.conectando || S.ocupado ? ' <span class="cm-fraco">…</span>' : ''}</div>
        ${motivo && !con ? `<div class="cm-aviso">${escH(motivo)} O "Testar sem aparelho" funciona assim mesmo.</div>` : ''}
        ${S.aviso ? `<div class="cm-aviso">${escH(S.aviso)}</div>` : ''}
        ${S.semBD ? `<div class="cm-aviso">Este navegador não deixou guardar o histórico. Exporte o CSV antes de fechar a tela.</div>` : ''}
      </div>
      ${S.pedido && !S.pedido.lugar ? htmlPedido() : ''}
      ${htmlRel()}${htmlCfg()}${htmlHist()}
      <div class="cm-colunas">
        <div class="cm-esq">
          <div class="card">${tabelaPagina(ref, amo, r)}</div>
          <div class="card">
            <div class="cm-centro">${escH(MODOS[modoAtual()] || 'SCI')}/${escH(cfg.iluminante1)}°</div>
            ${barra}
            <div class="cm-blocos">${blocoCor(ref, 'Padrão', 'padrao')}${blocoCor(amo, 'Amostra', 'amostra')}</div>
          </div>
        </div>
        <div class="cm-dir">
          <div class="card cm-aparelho">
            <div><div class="card-title">Aparelho</div><div class="cm-peq">${info}</div></div>
            <div><div class="card-title">Calibração</div><div class="cm-peq">${cal}</div></div>
            <div class="cm-peq" style="grid-column:1/-1">Botão do aparelho mede como:
              <label><input type="radio" name="cm-botao" value="amostra" ${cfg.botao === 'amostra' ? 'checked' : ''} onchange="Colorimetro.botaoComo(this.value)"> Amostra</label>
              <label><input type="radio" name="cm-botao" value="padrao" ${cfg.botao === 'padrao' ? 'checked' : ''} onchange="Colorimetro.botaoComo(this.value)"> Padrão</label>
              ${cfg.media > 1 ? `<span class="cm-fraco"> · média de ${cfg.media} leituras</span>` : ''}
            </div>
          </div>
          <div class="card"><canvas id="cm-espectro" class="cm-espectro"></canvas></div>
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
    desenharEspectro();
  }

  function desenharEspectro() {
    const cv = document.getElementById('cm-espectro');
    if (!cv) return;
    const css = getComputedStyle(document.documentElement);
    const tinta = (css.getPropertyValue('--ink-3') || '#6b6b6b').trim();
    const linha = (css.getPropertyValue('--line') || '#d4d0c5').trim();
    const texto = (css.getPropertyValue('--ink') || '#1a1a1a').trim();
    const dpr = window.devicePixelRatio || 1;
    const w = cv.clientWidth || 600, h = 220;
    cv.width = w * dpr; cv.height = h * dpr; cv.style.height = h + 'px';
    const g = cv.getContext('2d');
    g.scale(dpr, dpr);
    g.clearRect(0, 0, w, h);
    const my = 26, x0 = 44, y1 = h - 24, x1 = w - 16;
    g.font = '12px system-ui, sans-serif'; g.fillStyle = texto; g.textBaseline = 'middle';
    g.fillText('Refletância espectral', 6, 10);
    const series = [[S.padrao, '#2a9fd6'], [S.atual !== S.padrao ? S.atual : null, '#eb6834']].filter(([s]) => s && s.espectro && s.espectro.length);
    let maxv = 100;
    series.forEach(([s]) => { maxv = Math.max(maxv, Math.max(...s.espectro) * 1.05); });
    const X = wl => x0 + (wl - 400) / 300 * (x1 - x0);
    const Y = v => y1 - v / maxv * (y1 - my);
    g.font = '10px system-ui, sans-serif';
    for (let v = 0; v <= maxv; v += 20) {
      g.strokeStyle = linha; g.beginPath(); g.moveTo(x0, Y(v)); g.lineTo(x1, Y(v)); g.stroke();
      g.fillStyle = tinta; g.textAlign = 'right'; g.fillText(v + '%', x0 - 6, Y(v));
    }
    g.textAlign = 'center';
    for (let wl = 400; wl <= 700; wl += 50) g.fillText(wl + ' nm', X(wl), y1 + 12);
    series.forEach(([s, cor]) => {
      g.strokeStyle = cor; g.lineWidth = 2.5; g.beginPath();
      let primeiro = true;
      s.espectro.forEach((v, i) => {
        const wl = s.wave_start + 10 * i;
        if (wl < 400 || wl > 700) return;
        if (primeiro) { g.moveTo(X(wl), Y(v)); primeiro = false; } else g.lineTo(X(wl), Y(v));
      });
      g.stroke(); g.lineWidth = 1;
    });
    g.textAlign = 'left';
    series.forEach(([s, cor], k) => {
      const xl = x1 - 220 + 110 * k;
      g.strokeStyle = cor; g.lineWidth = 3; g.beginPath(); g.moveTo(xl, 10); g.lineTo(xl + 18, 10); g.stroke(); g.lineWidth = 1;
      g.fillStyle = texto; g.fillText(s.nome, xl + 24, 10);
    });
    if (!series.length) {
      g.fillStyle = tinta; g.textAlign = 'center';
      g.fillText('A curva aparece depois da primeira medição', (x0 + x1) / 2, (my + y1) / 2);
    }
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
    .cm-bloco{height:150px;border-radius:4px;display:flex;align-items:flex-end;justify-content:center;padding:10px;font-size:15px;cursor:pointer;border:1px solid var(--line)}
    .cm-bloco-vazio{background:repeating-linear-gradient(45deg,var(--line-2),var(--line-2) 8px,transparent 8px,transparent 16px);color:var(--ink-3)}
    .cm-aparelho{display:grid;grid-template-columns:1fr 1fr;gap:10px}
    .cm-peq{font-size:12px;line-height:1.6}
    .cm-peq label{margin-left:10px}
    .cm-espectro{width:100%;display:block}
    .cm-rolar{max-height:46vh;overflow:auto}
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
    .cm-av{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px;margin:10px 0}
    .cm-av-bloco{padding:8px 10px;background:var(--line-2);border-radius:2px}
    .cm-av-atual{outline:2px solid var(--accent)}
    .cm-av-rot{font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--ink-3)}
    .cm-av-num{font-size:24px;font-weight:700;margin:2px 0}
    .cm-fases{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
    .cm-fase{padding:4px 10px;border:1px solid var(--line);border-radius:2px;font-size:12px;color:var(--ink-3)}
    .cm-fase-atual{border-color:var(--accent);color:var(--ink);font-weight:600;outline:1px solid var(--accent)}
    .cm-fase-feita{color:var(--ok)}
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
    window.addEventListener('resize', () => { if (visivel()) desenharEspectro(); });
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
    aplicarCfg, resetarCfg, responder, avaliarOS, encerrarOS, relMudar, relSalvar, avTrocar, avIrPara, avReavaliar, avConcluirFase, avReavaliarFase,
    pagina: i => { S.pagina = i; desenhar(); },
    botaoComo: v => { cfg.botao = v === 'padrao' ? 'padrao' : 'amostra'; salvarCfg(); },
    simularBotao: () => { if (S.t && S.t.apertarBotao) S.t.apertarBotao(modoAtual()); },
    // para os testes: as contas, sem tela
    _contas: { carregarTabelas, espectroParaLab, deAb, de2000, de94, deCmc, labHex, comparar, arred, num, pacote, checksumOk }
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', instalar);
  else instalar();
})();
