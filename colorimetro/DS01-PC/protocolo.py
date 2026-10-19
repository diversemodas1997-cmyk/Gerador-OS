"""
Protocolo Bluetooth LE do colorímetro CHNSpec DS-01 (família ColorMeter).

Extraído do app Android oficial (com.hzcaipu.colormeter):
  - Serviço FFE0, característica FFE1 (escrita sem resposta + notificações)
  - Antes de cada comando envia-se o byte 0x01 para "acordar" o aparelho
  - Pacote: BB <cmd> <dados...> FF <checksum>, checksum = soma dos bytes % 256
"""
import asyncio
import json
import math
import os
import random
import time
from dataclasses import dataclass, field

SERVICE_UUID = "0000ffe0-0000-1000-8000-00805f9b34fb"
CHAR_UUID = "0000ffe1-0000-1000-8000-00805f9b34fb"

# Prefixos de nome que o app oficial aceita na busca
NOMES_CONHECIDOS = ("CM", "NJ", "GJ", "CH", "DS")

MODOS = {0: "SCI", 1: "SCE", 3: "M0", 4: "M1", 5: "M2", 6: "M3"}

AQUI = os.path.dirname(os.path.abspath(__file__))


# --------------------------------------------------------------------------- pacotes

def pacote(*bytes_):
    b = [int(x) & 0xFF for x in bytes_]
    return bytes(b + [sum(b) % 256])


def checksum_ok(data):
    return len(data) >= 2 and sum(data[:-1]) % 256 == data[-1]


def texto(buf):
    out = bytearray()
    for c in buf:
        if c == 0:
            break
        out.append(c)
    return out.decode("utf-8", "replace").strip()


def float32_le(data, offset):
    import struct
    return struct.unpack_from("<f", bytes(data), offset)[0]


# --------------------------------------------------------------------------- cor

class Colorimetria:
    """Espectro -> XYZ -> L*a*b* com as mesmas tabelas e fórmulas do app oficial."""

    def __init__(self, caminho=None):
        caminho = caminho or os.path.join(AQUI, "iluminantes.json")
        with open(caminho, encoding="utf-8") as f:
            t = json.load(f)
        self.wave_start = t["wave_start"]
        self.count = t["count"]
        self.iluminantes = t["illuminants"]

    def nomes(self):
        return list(self.iluminantes.keys())

    def completar(self, espectro, wave_start):
        """Alinha o espectro na grade 360-780 nm (43 pontos), como fillSpectralArray do app."""
        if len(espectro) == self.count:
            return list(espectro)
        ini = abs(self.wave_start - wave_start) // 10
        dados = [0.0] * ini + list(espectro)
        dados += [0.0] * max(0, self.count - len(dados))
        return dados[: self.count]

    def xyz(self, espectro43, iluminante="D65/10"):
        w = self.iluminantes[iluminante]["weights"]
        return [sum(w[c][i] * espectro43[i] for i in range(self.count)) for c in range(3)]

    def lab(self, xyz, iluminante="D65/10"):
        # Mesma fórmula (e constantes) de XYZtoCIE_LabCH do app
        xn, yn, zn = self.iluminantes[iluminante]["white"]

        def f(t):
            return t ** 0.3333333 if t > 0.008856 else 7.787 * t + 0.138

        ty = xyz[1] / yn
        L = 116 * ty ** 0.3333333 - 16 if ty > 0.008856 else 903.3 * ty
        a = 500 * (f(xyz[0] / xn) - f(ty))
        b = 200 * (f(ty) - f(xyz[2] / zn))
        return max(L, 0.0), a, b

    def espectro_para_lab(self, espectro, wave_start, iluminante="D65/10"):
        e = self.completar(espectro, wave_start)
        return self.lab(self.xyz(e, iluminante), iluminante)


def lch(L, a, b):
    C = math.hypot(a, b)
    h = math.degrees(math.atan2(b, a)) % 360
    return C, h


def delta_e_ab(lab1, lab2):
    return math.dist(lab1, lab2)


def delta_e_2000(lab1, lab2, kL=1.0, kC=1.0, kH=1.0):
    L1, a1, b1 = lab1
    L2, a2, b2 = lab2
    C1, C2 = math.hypot(a1, b1), math.hypot(a2, b2)
    Cm = (C1 + C2) / 2
    G = 0.5 * (1 - math.sqrt(Cm ** 7 / (Cm ** 7 + 25 ** 7)))
    a1p, a2p = (1 + G) * a1, (1 + G) * a2
    C1p, C2p = math.hypot(a1p, b1), math.hypot(a2p, b2)
    h1p = math.degrees(math.atan2(b1, a1p)) % 360 if C1p else 0.0
    h2p = math.degrees(math.atan2(b2, a2p)) % 360 if C2p else 0.0
    dLp = L2 - L1
    dCp = C2p - C1p
    if C1p * C2p == 0:
        dhp = 0.0
    elif abs(h2p - h1p) <= 180:
        dhp = h2p - h1p
    elif h2p - h1p > 180:
        dhp = h2p - h1p - 360
    else:
        dhp = h2p - h1p + 360
    dHp = 2 * math.sqrt(C1p * C2p) * math.sin(math.radians(dhp / 2))
    Lm = (L1 + L2) / 2
    Cmp = (C1p + C2p) / 2
    if C1p * C2p == 0:
        hm = h1p + h2p
    elif abs(h1p - h2p) <= 180:
        hm = (h1p + h2p) / 2
    elif h1p + h2p < 360:
        hm = (h1p + h2p + 360) / 2
    else:
        hm = (h1p + h2p - 360) / 2
    T = (1 - 0.17 * math.cos(math.radians(hm - 30)) + 0.24 * math.cos(math.radians(2 * hm))
         + 0.32 * math.cos(math.radians(3 * hm + 6)) - 0.20 * math.cos(math.radians(4 * hm - 63)))
    dtheta = 30 * math.exp(-(((hm - 275) / 25) ** 2))
    Rc = 2 * math.sqrt(Cmp ** 7 / (Cmp ** 7 + 25 ** 7))
    Sl = 1 + 0.015 * (Lm - 50) ** 2 / math.sqrt(20 + (Lm - 50) ** 2)
    Sc = 1 + 0.045 * Cmp
    Sh = 1 + 0.015 * Cmp * T
    Rt = -math.sin(math.radians(2 * dtheta)) * Rc
    return math.sqrt((dLp / (kL * Sl)) ** 2 + (dCp / (kC * Sc)) ** 2 + (dHp / (kH * Sh)) ** 2
                     + Rt * (dCp / (kC * Sc)) * (dHp / (kH * Sh)))


def delta_e_94(ref, amo, kL=1.0, kC=1.0, kH=1.0):
    L1, a1, b1 = ref
    L2, a2, b2 = amo
    C1, C2 = math.hypot(a1, b1), math.hypot(a2, b2)
    dL, dC = L1 - L2, C1 - C2
    dH2 = max(0.0, (a1 - a2) ** 2 + (b1 - b2) ** 2 - dC ** 2)
    Sc, Sh = 1 + 0.045 * C1, 1 + 0.015 * C1
    return math.sqrt((dL / kL) ** 2 + (dC / (kC * Sc)) ** 2 + dH2 / (kH * Sh) ** 2)


def delta_e_cmc(ref, amo, l=2.0, c=1.0):
    L1, a1, b1 = ref
    L2, a2, b2 = amo
    C1, C2 = math.hypot(a1, b1), math.hypot(a2, b2)
    dL, dC = L1 - L2, C1 - C2
    dH2 = max(0.0, (a1 - a2) ** 2 + (b1 - b2) ** 2 - dC ** 2)
    h1 = math.degrees(math.atan2(b1, a1)) % 360
    F = math.sqrt(C1 ** 4 / (C1 ** 4 + 1900))
    T = (0.56 + abs(0.2 * math.cos(math.radians(h1 + 168))) if 164 <= h1 <= 345
         else 0.36 + abs(0.4 * math.cos(math.radians(h1 + 35))))
    Sl = 0.511 if L1 < 16 else 0.040975 * L1 / (1 + 0.01765 * L1)
    Sc = 0.0638 * C1 / (1 + 0.0131 * C1) + 0.638
    Sh = Sc * (F * T + 1 - F)
    return math.sqrt((dL / (l * Sl)) ** 2 + (dC / (c * Sc)) ** 2 + dH2 / Sh ** 2)


def componentes_00(ref, amo, kL=1.0, kC=1.0, kH=1.0):
    """dL', dC', dH' ponderados do CIEDE2000 (para tolerância por componente)."""
    L1, a1, b1 = ref
    L2, a2, b2 = amo
    C1, C2 = math.hypot(a1, b1), math.hypot(a2, b2)
    Cm = (C1 + C2) / 2
    G = 0.5 * (1 - math.sqrt(Cm ** 7 / (Cm ** 7 + 25 ** 7)))
    a1p, a2p = (1 + G) * a1, (1 + G) * a2
    C1p, C2p = math.hypot(a1p, b1), math.hypot(a2p, b2)
    h1p = math.degrees(math.atan2(b1, a1p)) % 360
    h2p = math.degrees(math.atan2(b2, a2p)) % 360
    dh = h2p - h1p
    if dh > 180:
        dh -= 360
    elif dh < -180:
        dh += 360
    dHp = 2 * math.sqrt(C1p * C2p) * math.sin(math.radians(dh / 2))
    Lm, Cmp = (L1 + L2) / 2, (C1p + C2p) / 2
    hm = (h1p + h2p) / 2 if abs(h1p - h2p) <= 180 else (h1p + h2p + 360) / 2
    T = (1 - 0.17 * math.cos(math.radians(hm - 30)) + 0.24 * math.cos(math.radians(2 * hm))
         + 0.32 * math.cos(math.radians(3 * hm + 6)) - 0.20 * math.cos(math.radians(4 * hm - 63)))
    Sl = 1 + 0.015 * (Lm - 50) ** 2 / math.sqrt(20 + (Lm - 50) ** 2)
    return (L2 - L1) / (kL * Sl), (C2p - C1p) / (kC * (1 + 0.045 * Cmp)), dHp / (kH * (1 + 0.015 * Cmp * T))


def lab_para_hex(L, a, b):
    """Cor aproximada para exibir na tela (Lab D65 -> sRGB)."""
    fy = (L + 16) / 116
    fx, fz = fy + a / 500, fy - b / 200

    def inv(t):
        return t ** 3 if t ** 3 > 0.008856 else (t - 16 / 116) / 7.787

    X, Y, Z = 0.95047 * inv(fx), 1.0 * inv(fy), 1.08883 * inv(fz)
    r = 3.2406 * X - 1.5372 * Y - 0.4986 * Z
    g = -0.9689 * X + 1.8758 * Y + 0.0415 * Z
    bl = 0.0557 * X - 0.2040 * Y + 1.0570 * Z

    def gamma(c):
        c = max(0.0, min(1.0, c))
        return 12.92 * c if c <= 0.0031308 else 1.055 * c ** (1 / 2.4) - 0.055

    return "#%02x%02x%02x" % tuple(round(gamma(c) * 255) for c in (r, g, bl))


# --------------------------------------------------------------------------- resultados

@dataclass
class InfoAparelho:
    codigo: int
    modelo: str
    serie: str
    firmware: str
    hardware: str

    def protocolo_novo(self):
        # Regra do app: ColorMeter SE (código 776) com firmware >= 20210420
        try:
            data = int(self.firmware.split(".")[-1])
        except ValueError:
            return False
        return self.codigo == 776 and data >= 20210420

    def modo_padrao(self):
        return 3 if self.codigo in (778, 779) else 0


@dataclass
class EstadoCalibracao:
    branco_ok: bool
    branco_data: int
    preto_ok: bool
    preto_data: int


@dataclass
class Medicao:
    modo: int
    wave_start: int = 400
    intervalo: int = 10
    espectro: list = field(default_factory=list)   # refletância em %
    lab_aparelho: tuple = None                      # quando o aparelho envia Lab pronto
    horario: float = field(default_factory=time.time)


class ErroAparelho(Exception):
    pass


# --------------------------------------------------------------------------- cliente

class DS01:
    """Fala o protocolo do aparelho sobre um 'transporte' (Bluetooth real ou simulado)."""

    def __init__(self, transporte, ao_medir_no_botao=None, log=print):
        self.t = transporte
        self.t.ao_receber = self._receber
        self.ao_medir_no_botao = ao_medir_no_botao
        self.log = log
        self.info = None
        self._buf = bytearray()
        self._evento = asyncio.Event()
        self._ocupado = asyncio.Lock()
        self._esperando = False

    # ---- baixo nível
    def _receber(self, data):
        data = bytes(data)
        if self._esperando:
            self._buf += data
            self._evento.set()
            return
        # Notificação espontânea: botão de medir apertado no aparelho
        if len(data) >= 10 and data[0] == 0xBB and data[1] == 0x01 and checksum_ok(data[:10]):
            if self.ao_medir_no_botao:
                self.ao_medir_no_botao(data[2], data[3])
        else:
            self.log(f"recebido fora de hora: {data.hex(' ')}")

    async def _escrever(self, data):
        for i in range(0, len(data), 20):
            await self.t.escrever(data[i:i + 20])
            await asyncio.sleep(0.01)

    def _resposta(self, cmd):
        """Recorta do buffer a resposta que começa em BB <cmd>, descartando sobras antes dela
        (pela ponte do celular, ecos atrasados de comandos anteriores podem chegar misturados)."""
        i = self._buf.find(bytes([0xBB, cmd]))
        return bytes(self._buf[i:]) if i >= 0 else b""

    async def comando(self, quadro, tamanho=10, timeout=3.0):
        # Leituras podem ser repetidas sem efeito colateral; medir e calibrar não.
        tentativas = 1 if quadro[1] in (0x01, 0x10, 0x11) else 2
        for tentativa in range(1, tentativas + 1):
            try:
                return await self._comando(quadro, tamanho, timeout)
            except ErroAparelho as e:
                self.log(f"{e} (tentativa {tentativa}); bytes recebidos: {bytes(self._buf).hex(' ') or 'nenhum'}")
                if tentativa == tentativas:
                    raise
                await asyncio.sleep(0.3)

    async def _comando(self, quadro, tamanho, timeout):
        cmd = quadro[1]
        async with self._ocupado:
            self._buf = bytearray()
            self._esperando = True
            try:
                await self.t.escrever(b"\x01")          # acordar
                await asyncio.sleep(0.1)
                self._buf = bytearray()
                await self._escrever(quadro)
                fim = time.monotonic() + timeout
                while len(self._resposta(cmd)) < tamanho:
                    resto = fim - time.monotonic()
                    if resto <= 0:
                        break
                    self._evento.clear()
                    try:
                        await asyncio.wait_for(self._evento.wait(), resto)
                    except asyncio.TimeoutError:
                        break
                await asyncio.sleep(0.05)               # sobras do último pacote
                data = self._resposta(cmd)
            finally:
                self._esperando = False
        if len(data) < tamanho:
            raise ErroAparelho(f"resposta incompleta ({len(data)}/{tamanho} bytes) ao comando {cmd:#04x}")
        if not any(checksum_ok(data[:n]) for n in {len(data), tamanho, 10, 20, 200} if n <= len(data)):
            raise ErroAparelho(f"checksum inválido na resposta ao comando {cmd:#04x}")
        return data

    # ---- comandos
    async def ler_info(self):
        d = await self.comando(pacote(0xBB, 0x12, 0x01, 0, 0, 0, 0, 0, 0xFF), 200)
        self.info = InfoAparelho(
            codigo=d[5] + (d[6] << 8),
            modelo=texto(d[37:67]),
            serie=texto(d[67:97]),
            firmware=texto(d[97:127]),
            hardware=texto(d[127:157]),
        )
        return self.info

    async def ler_calibracao(self):
        d = await self.comando(pacote(0xBB, 0x1E, 0, 0, 0, 0, 0, 0, 0xFF), 10)
        d = d + bytes(max(0, 12 - len(d)))
        return EstadoCalibracao(
            branco_ok=d[2] == 1,
            branco_data=int.from_bytes(d[3:7], "little"),
            preto_ok=d[7] == 1,
            preto_data=int.from_bytes(d[8:12], "little"),
        )

    async def calibrar(self, tipo):
        cmd = 0x10 if tipo == "preto" else 0x11
        d = await self.comando(pacote(0xBB, cmd, 1, 0, 0, 0, 0, 0, 0xFF), 10, timeout=12)
        return d[2] == 0

    async def sincronizar_relogio(self):
        ts = int(time.time()).to_bytes(4, "little")
        await self.comando(pacote(0xBB, 0x14, *ts, 0, 0, 0xFF), 10)

    async def medir(self, modo=None):
        if modo is None:
            modo = self.info.modo_padrao() if self.info else 0
        aleatorio = [random.randint(0, 255) for _ in range(4)]
        d = await self.comando(pacote(0xBB, 0x01, modo, *aleatorio, 0, 0xFF), 10, timeout=6)
        if d[3] != 0:
            raise ErroAparelho(f"o aparelho retornou erro {d[3]} na medição")
        return await self.baixar_medicao(d[2])

    async def baixar_medicao(self, modo):
        cmd = 0x05 if (self.info and self.info.protocolo_novo()) else 0x02
        d = await self.comando(pacote(0xBB, cmd, modo + 0x10, 0, 0, 0, 0, 0, 0xFF), 200)
        m = Medicao(modo=d[2] - 0x10, wave_start=(d[4] << 8) + d[5], intervalo=d[6])
        tipo, n = d[3], d[7]
        if tipo == 0:
            m.espectro = [max(0.0, float32_le(d, 8 + 4 * i)) for i in range(n)]
        else:
            m.lab_aparelho = tuple(float32_le(d, o) for o in (184, 188, 192))
        return m


# --------------------------------------------------------------------------- transportes

class TransporteBLE:
    """Conexão real via Bluetooth LE (biblioteca bleak)."""

    def __init__(self):
        self.cliente = None
        self.ao_receber = None
        self.ao_desconectar = None

    @staticmethod
    async def procurar(segundos=8):
        from bleak import BleakScanner
        achados = await BleakScanner.discover(timeout=segundos, return_adv=True)
        lista = []
        for dev, adv in achados.values():
            nome = adv.local_name or dev.name or ""
            uuids = [u.lower() for u in (adv.service_uuids or [])]
            provavel = nome.upper().startswith(NOMES_CONHECIDOS) or SERVICE_UUID in uuids
            if nome or provavel:
                lista.append({"endereco": dev.address, "nome": nome or "(sem nome)",
                              "rssi": adv.rssi, "provavel": provavel})
        lista.sort(key=lambda x: (not x["provavel"], -(x["rssi"] or -999)))
        return lista

    async def conectar(self, endereco):
        from bleak import BleakClient

        def desconectou(_):
            if self.ao_desconectar:
                self.ao_desconectar()

        # No Windows a 1ª conexão costuma cair com "Operação anulada" logo após conectar;
        # limitar a descoberta ao serviço FFE0 e tentar de novo resolve.
        for tentativa in range(1, 5):
            self.cliente = BleakClient(endereco, disconnected_callback=desconectou,
                                       services=[SERVICE_UUID])
            try:
                await self.cliente.connect(timeout=20)
                await self.cliente.start_notify(CHAR_UUID, lambda _c, data: self.ao_receber(data))
                return
            except Exception:
                try:
                    await self.cliente.disconnect()
                except Exception:
                    pass
                if tentativa == 4:
                    raise
                await asyncio.sleep(2)

    async def escrever(self, data):
        await self.cliente.write_gatt_char(CHAR_UUID, bytes(data), response=False)

    async def reassinar(self):
        try:
            await self.cliente.stop_notify(CHAR_UUID)
        except Exception:
            pass
        await self.cliente.start_notify(CHAR_UUID, lambda _c, data: self.ao_receber(data))

    async def desconectar(self):
        if self.cliente:
            try:
                await self.cliente.stop_notify(CHAR_UUID)
            except Exception:
                pass
            await self.cliente.disconnect()


class TransporteCelular:
    """Usa o Bluetooth do celular como ponte: o Chrome do Android abre ponte.html,
    conecta no DS-01 por Web Bluetooth e repassa os bytes por WebSocket."""

    PORTA = 8765
    ENDERECO = "celular"
    _servidor = None
    _clientes = {}         # página da ponte -> {"conectado": bool, "nome": str}
    _ws = None
    _ble_conectado = False
    _nome = None
    _mudou = None
    _ativo = None          # transporte que recebe as notificações

    def __init__(self):
        self.ao_receber = None
        self.ao_desconectar = None

    @staticmethod
    def enderecos():
        import socket
        ips = []
        try:
            for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
                ip = info[4][0]
                if ip.startswith(("192.168.", "10.")) and ip not in ips:
                    ips.append(ip)
        except OSError:
            pass
        return [f"http://{ip}:{TransporteCelular.PORTA}" for ip in ips]

    @classmethod
    async def iniciar_servidor(cls):
        if cls._servidor:
            return
        from websockets.asyncio.server import serve
        cls._mudou = asyncio.Event()
        with open(os.path.join(AQUI, "ponte.html"), encoding="utf-8") as f:
            pagina = f.read()

        def http(conexao, pedido):
            if pedido.path.split("?")[0] == "/ws":
                return None
            resp = conexao.respond(200, pagina)
            del resp.headers["Content-Type"]
            resp.headers["Content-Type"] = "text/html; charset=utf-8"
            return resp

        cls._servidor = await serve(cls._atender, "0.0.0.0", cls.PORTA, process_request=http,
                                    ping_interval=10, ping_timeout=10, max_size=2 ** 16)

    @classmethod
    def _sinalizar(cls):
        cls._mudou.set()
        cls._mudou = asyncio.Event()

    @classmethod
    def _recalcular(cls):
        """Várias páginas podem estar abertas (celular, outra aba...): usa a que está
        ligada ao colorímetro; sem nenhuma ligada, a aberta mais recentemente."""
        estava, ws_antes = cls._ble_conectado, cls._ws
        ligadas = [w for w, e in cls._clientes.items() if e["conectado"]]
        if ligadas:
            cls._ws = ws_antes if ws_antes in ligadas else ligadas[-1]
        else:
            cls._ws = next(reversed(cls._clientes), None)
        e = cls._clientes.get(cls._ws, {})
        cls._ble_conectado, cls._nome = bool(e.get("conectado")), e.get("nome")
        if estava and (not cls._ble_conectado or cls._ws is not ws_antes):
            cls._caiu()
        cls._sinalizar()

    @classmethod
    async def _atender(cls, ws):
        cls._clientes[ws] = {"conectado": False, "nome": None}
        cls._recalcular()
        try:
            async for msg in ws:
                if isinstance(msg, bytes):
                    if ws is cls._ws and cls._ativo and cls._ativo.ao_receber:
                        cls._ativo.ao_receber(msg)
                    continue
                m = json.loads(msg)
                if m.get("tipo") == "estado":
                    cls._clientes[ws] = {"conectado": bool(m.get("conectado")), "nome": m.get("nome")}
                    cls._recalcular()
        finally:
            cls._clientes.pop(ws, None)
            cls._recalcular()

    @classmethod
    def _caiu(cls):
        t = cls._ativo
        if t and t.ao_desconectar:
            t.ao_desconectar()

    @classmethod
    async def _esperar(cls, condicao, segundos):
        fim = time.monotonic() + segundos
        while not condicao():
            resto = fim - time.monotonic()
            if resto <= 0:
                return False
            try:
                await asyncio.wait_for(cls._mudou.wait(), resto)
            except asyncio.TimeoutError:
                return condicao()
        return True

    @classmethod
    async def procurar(cls, segundos=8):
        await cls.iniciar_servidor()
        if cls._ws is None:
            nome = "📱 Pelo celular (abra a ponte no Chrome do celular)"
        elif cls._ble_conectado:
            nome = f"📱 {cls._nome} pelo celular"
        else:
            nome = "📱 Pelo celular (toque em Conectar no celular)"
        return [{"endereco": cls.ENDERECO, "nome": nome, "rssi": None, "provavel": True}]

    async def conectar(self, endereco):
        cls = TransporteCelular
        await cls.iniciar_servidor()
        cls._ativo = self
        if not await cls._esperar(lambda: cls._ws is not None, 60):
            raise ErroAparelho("o celular não abriu a página da ponte. No Chrome do celular abra: "
                               + " ou ".join(cls.enderecos()))
        if not cls._ble_conectado:
            for ws in list(cls._clientes):      # quem já escolheu o aparelho reconecta sozinho
                try:
                    await ws.send(json.dumps({"tipo": "conectar"}))
                except Exception:
                    pass
            if not await cls._esperar(lambda: cls._ble_conectado, 90):
                raise ErroAparelho("o celular não conectou no colorímetro. "
                                   "Toque em 'Conectar ao colorímetro' na página da ponte.")

    async def escrever(self, data):
        ws = TransporteCelular._ws
        if ws is None or not TransporteCelular._ble_conectado:
            raise ErroAparelho("a ponte do celular está desconectada")
        await ws.send(bytes(data))

    async def reassinar(self):
        pass

    async def desconectar(self):
        cls = TransporteCelular
        if cls._ativo is self:
            cls._ativo = None
        if cls._ws is not None:
            try:
                await cls._ws.send(json.dumps({"tipo": "desconectar"}))
            except Exception:
                pass


class TransporteSimulado:
    """Aparelho falso que responde ao mesmo protocolo — para testar sem o DS-01."""

    def __init__(self):
        self.ao_receber = None
        self.ao_desconectar = None
        self.ultimo_modo = 0
        self.cal = {0x10: int(time.time()) - 3600, 0x11: int(time.time()) - 3600}

    @staticmethod
    async def procurar(segundos=1):
        await asyncio.sleep(0.5)
        return [{"endereco": "00:00:00:00:D5:01", "nome": "CM-DS01 (simulado)", "rssi": -50, "provavel": True}]

    async def conectar(self, endereco):
        self.loop = asyncio.get_running_loop()
        await asyncio.sleep(0.3)

    async def desconectar(self):
        pass

    async def reassinar(self):
        pass

    def _responder(self, data, atraso=0.05):
        async def enviar():
            await asyncio.sleep(atraso)
            for i in range(0, len(data), 20):          # simula fragmentação BLE
                self.ao_receber(data[i:i + 20])
                await asyncio.sleep(0.005)
        asyncio.get_running_loop().create_task(enviar())

    @staticmethod
    def _fechar(arr):
        arr = bytearray(arr)
        arr[-1] = sum(arr[:-1]) % 256
        return bytes(arr)

    def apertar_botao(self, modo=0):
        self.ultimo_modo = modo
        quadro = pacote(0xBB, 0x01, modo, 0, 0, 0, 0, 0, 0xFF)
        self.loop.call_soon_threadsafe(self._responder, quadro)

    async def escrever(self, data):
        if len(data) < 9 or data[0] != 0xBB:
            return
        cmd = data[1]
        if cmd == 0x12:
            r = bytearray(200)
            r[0], r[1] = 0xBB, 0x12
            r[5], r[6] = 774 & 0xFF, 774 >> 8
            for off, s in ((37, "DS-01"), (67, "SIM0000001"), (97, "V1.0.20240101"), (127, "H1.0")):
                r[off:off + len(s)] = s.encode()
            r[198] = 0xFF
            self._responder(self._fechar(r))
        elif cmd == 0x1E:
            r = bytearray(20)
            r[0], r[1], r[2] = 0xBB, 0x1E, 1
            r[3:7] = self.cal[0x11].to_bytes(4, "little")
            r[7] = 1
            r[8:12] = self.cal[0x10].to_bytes(4, "little")
            r[18] = 0xFF
            self._responder(self._fechar(r))
        elif cmd in (0x10, 0x11):
            self.cal[cmd] = int(time.time())
            self._responder(pacote(0xBB, cmd, 0, 0, 0, 0, 0, 0, 0xFF), atraso=1.5)
        elif cmd == 0x14:
            self._responder(pacote(0xBB, 0x14, 0, 0, 0, 0, 0, 0, 0xFF))
        elif cmd == 0x01:
            self.ultimo_modo = data[2]
            self._responder(pacote(0xBB, 0x01, data[2], 0, 0x01, 0x90, 10, 31, 0xFF), atraso=0.8)
        elif cmd in (0x02, 0x05):
            import struct
            r = bytearray(200)
            r[0], r[1], r[2], r[3] = 0xBB, cmd, data[2], 0
            r[4], r[5], r[6], r[7] = 0x01, 0x90, 10, 31          # 400 nm, 10 nm, 31 pontos
            c = random.uniform(450, 650)
            base, amp = random.uniform(5, 30), random.uniform(20, 60)
            for i in range(31):
                wl = 400 + 10 * i
                v = base + amp * math.exp(-((wl - c) / random.uniform(50, 90)) ** 2)
                struct.pack_into("<f", r, 8 + 4 * i, v)
            r[198] = 0xFF
            self._responder(self._fechar(r))
