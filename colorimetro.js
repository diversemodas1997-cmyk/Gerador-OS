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
    pedido: null, mostrarCfg: false, mostrarHist: false, histDias: [], histDia: '', histLidas: [], semBD: false
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
  const refDe = l => l.refId ? S.leituras.find(x => x.id === l.refId) || null : null;

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
  function perguntar(pedido) { S.pedido = pedido; desenhar(); }
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
    S.ocupado = true;
    status('Medindo…');
    try { chegou(await S.dev.medir(modoAtual()), tipo); } catch (e) { erroDoAparelho(e); }
  }

  // O botão do próprio DS-01 foi apertado: a medição já foi feita, só baixar.
  function aoBotao(modo) {
    if (S.ocupado || !S.dev) return;
    let tipo = cfg.botao;
    if (tipo === 'amostra' && !S.padrao && !S.acumulando.length) tipo = 'padrao';
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
    S.cont[tipo]++;
    const prefixo = tipo === 'padrao' ? cfg.nome_padrao : cfg.nome_amostra;
    const l = novaLeitura({ nome: `${prefixo} ${S.cont[tipo]}`, tipo, espectro: esp, wave_start: medidas[0].wave_start,
      lab_aparelho: labAp, horario: new Date() });
    if (tipo === 'amostra' && S.padrao) l.refId = S.padrao.id;
    S.leituras.unshift(l);
    if (tipo === 'padrao') S.padrao = l;
    S.atual = l;
    guardar(l);
    status(`${l.nome} medido às ${horaTxt(l.horario)}`);
  }

  // Cada medição vai para o histórico deste navegador na hora, como o CSV do dia do original.
  function registroDe(l) {
    const ref = refDe(l);
    return { id: l.id, dia: diaIso(l.horario), horario: l.horario.toISOString(), nome: l.nome, tipo: l.tipo,
      espectro: l.espectro, wave_start: l.wave_start, lab_aparelho: l.lab_aparelho,
      refId: l.refId, refNome: ref ? ref.nome : '', iluminante: cfg.iluminante1 };
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
    perguntar({ titulo: 'Limpar a lista', texto: 'Apagar a lista da tela? O histórico do dia continua guardado.', sim: 'Limpar',
      executar: () => { S.leituras = []; S.padrao = null; S.atual = null; S.cont = { padrao: 0, amostra: 0 }; desenhar(); } });
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
    l.refId = r.refId; l.refNome = r.refNome;
    return l;
  }
  async function abrirHistorico(manter) {
    try {
      await carregarTabelas();
      const todos = await BD.todos();
      S.histDias = [...new Set(todos.map(r => r.dia))].sort().reverse();
      if (!manter || !S.histDias.includes(S.histDia)) S.histDia = S.histDias[0] || '';
      S.mostrarHist = true;
      await carregarDia(S.histDia);
    } catch (e) {
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
    S.mostrarHist = false;
    status(`${r.nome} de ${dataTxt(r.horario).slice(0, 5)} ${horaTxt(r.horario).slice(0, 5)} virou o padrão. Agora meça a amostra`);
  }
  function histExportar() {
    if (!S.histLidas.length) return;
    const porId = {};
    S.histLidas.forEach(l => { porId[l.id] = l; });
    baixarCsv(`medicoes_${S.histDia}.csv`, [cabecalhoCsv(), ...S.histLidas.map(l => linhaCsv(l, porId[l.refId]))]);
  }

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
    S.mostrarCfg = false;
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
    if (!S.mostrarCfg) return '';
    const ils = T ? Object.keys(T.illuminants) : [cfg.iluminante1, cfg.iluminante2];
    const opt = (lista, atual) => lista.map(x => `<option value="${escH(x[0])}" ${String(atual) === String(x[0]) ? 'selected' : ''}>${escH(x[1])}</option>`).join('');
    const campo = (id, rot, val, larg) => `<label class="cm-campo" style="${larg ? 'max-width:' + larg : ''}"><span>${rot}</span><input id="${id}" value="${escH(numTol(val))}" inputmode="decimal"></label>`;
    const tolRot = { dL: 'dL* ±', da: 'da* ±', db: 'db* ±', 'dE*ab': 'dE*ab ≤', dC: 'dC* ±', dH: 'dH* ±', "dL'": "dL' ±", "dC'": "dC' ±", "dH'": "dH' ±", dE00: 'dE00 ≤', dEcmc: 'dEcmc ≤', dE94: 'dE94 ≤' };
    return `<div class="card cm-cfg">
      <div class="card-title">Configurações do colorímetro <span class="cm-fraco">(valem neste computador)</span></div>
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
    if (!S.mostrarHist) return '';
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
      <div class="card-title">Histórico deste computador</div>
      ${S.histDias.length ? `<div class="cm-botoes">
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
      ${S.pedido ? `<div class="card cm-pedido"><div class="card-title">${escH(S.pedido.titulo)}</div>
        <div>${escH(S.pedido.texto)}</div>
        ${S.pedido.campo != null ? `<input id="cm-pedido-valor" class="cm-pedido-campo" value="${escH(S.pedido.campo)}" onkeydown="if(event.key==='Enter')Colorimetro.responder(true)">` : ''}
        <div class="cm-botoes"><button type="button" class="btn primary" onclick="Colorimetro.responder(true)">${escH(S.pedido.sim)}</button>
        <button type="button" class="btn" onclick="Colorimetro.responder(false)">Cancelar</button></div></div>` : ''}
      ${htmlCfg()}${htmlHist()}
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
                <button type="button" class="btn small" onclick="Colorimetro.abrirHistorico()">Histórico</button>
                <button type="button" class="btn small" onclick="Colorimetro.exportar()">Exportar CSV</button>
                <label class="btn small" title="Os arquivos medicoes_AAAA-MM-DD.csv do programa antigo">Importar CSV
                  <input type="file" accept=".csv,text/csv" hidden onchange="Colorimetro.importar(this.files[0]); this.value=''"></label>
                <button type="button" class="btn small" onclick="Colorimetro.abrirCfg()">Configurações</button>
                <button type="button" class="btn small ghost" onclick="Colorimetro.limpar()">Limpar lista</button>
              </div>
            </div>
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
    .cm-cfg,.cm-hist,.cm-pedido{margin-top:12px}
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
      if (!visivel() || !S.dev || e.ctrlKey || e.altKey || e.metaKey || e.shiftKey) return;
      if (e.key === 'F5' || e.key === 'F6') {
        e.preventDefault();
        medir(e.key === 'F6' ? 'padrao' : 'amostra');
      }
    });
  }

  window.Colorimetro = {
    conectar, desconectar, calibrar, medir, selecionar, usarPadrao, renomear, excluir, limpar, exportar, importar,
    abrirHistorico: () => abrirHistorico(false), carregarDia: d => carregarDia(d).catch(e => { S.aviso = e.message; desenhar(); }),
    histUsarPadrao, histExportar, fecharHist: () => { S.mostrarHist = false; desenhar(); },
    abrirCfg: () => { S.mostrarCfg = true; S.mostrarHist = false; desenhar(); }, fecharCfg: () => { S.mostrarCfg = false; desenhar(); },
    aplicarCfg, resetarCfg, responder,
    pagina: i => { S.pagina = i; desenhar(); },
    botaoComo: v => { cfg.botao = v === 'padrao' ? 'padrao' : 'amostra'; salvarCfg(); },
    simularBotao: () => { if (S.t && S.t.apertarBotao) S.t.apertarBotao(modoAtual()); },
    // para os testes: as contas, sem tela
    _contas: { carregarTabelas, espectroParaLab, deAb, de2000, de94, deCmc, labHex, comparar, arred, num, pacote, checksumOk }
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', instalar);
  else instalar();
})();
