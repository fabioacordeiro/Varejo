# Agora para rodar cole o prompt abaixo
# Gravar cliques	python mouse_macro.py --record C:\Fabio\Desenvolvimento\Varejo\Docusign\clicks.json
# Ver o conteúdo gravado	python mouse_macro.py --show C:\Fabio\Desenvolvimento\Varejo\Docusign\clicks.json
# Reproduzir os cliques	python mouse_macro.py --play C:\Fabio\Desenvolvimento\Varejo\Docusign\clicks.json

import time
import json
import argparse
from pathlib import Path

from pynput import mouse, keyboard
import pyautogui

# -------------- Configurações padrão --------------
STOP_KEY = keyboard.Key.f12          # tecla para encerrar gravação
DEFAULT_MIN_MOVE_PX = 8              # deslocamento mínimo entre movimentos gravados
DEFAULT_MIN_MOVE_INTERVAL = 0.02     # intervalo mínimo (s) entre gravações de movimento
pyautogui.FAILSAFE = True            # mover para (0,0) aborta playback

# -------------- Normalização de teclas --------------
# Mapeia nomes do pynput -> nomes aceitos pelo pyautogui
SPECIAL_KEY_MAP = {
    'Key.alt': 'alt', 'Key.alt_l': 'alt', 'Key.alt_r': 'alt',
    'Key.cmd': 'win', 'Key.cmd_l': 'win', 'Key.cmd_r': 'win',
    'Key.ctrl': 'ctrl', 'Key.ctrl_l': 'ctrl', 'Key.ctrl_r': 'ctrl',
    'Key.shift': 'shift', 'Key.shift_l': 'shift', 'Key.shift_r': 'shift',
    'Key.enter': 'enter', 'Key.esc': 'esc', 'Key.tab': 'tab',
    'Key.space': 'space', 'Key.backspace': 'backspace', 'Key.delete': 'delete',
    'Key.insert': 'insert', 'Key.home': 'home', 'Key.end': 'end',
    'Key.page_up': 'pageup', 'Key.page_down': 'pagedown',
    'Key.left': 'left', 'Key.right': 'right', 'Key.up': 'up', 'Key.down': 'down',
    'Key.media_volume_up': 'volumeup', 'Key.media_volume_down': 'volumedown',
    'Key.media_volume_mute': 'volumemute',
    'Key.f1': 'f1', 'Key.f2': 'f2', 'Key.f3': 'f3', 'Key.f4': 'f4',
    'Key.f5': 'f5', 'Key.f6': 'f6', 'Key.f7': 'f7', 'Key.f8': 'f8',
    'Key.f9': 'f9', 'Key.f10': 'f10', 'Key.f11': 'f11', 'Key.f12': 'f12',
}

def norm_key(k) -> str | None:
    """Converte uma tecla do pynput em nome aceito pelo pyautogui."""
    try:
        # Caracter (a, b, 1, etc.)
        if hasattr(k, 'char') and k.char is not None:
            return k.char
        # Tecla especial
        name = str(k)
        return SPECIAL_KEY_MAP.get(name, None)
    except Exception:
        return None

def norm_button(btn) -> str:
    """Converte botão do mouse do pynput para string ('left','right','middle')."""
    s = str(btn)
    if 'left' in s:
        return 'left'
    if 'right' in s:
        return 'right'
    if 'middle' in s:
        return 'middle'
    return 'left'

# -------------- Gravação --------------
def record(output_file: str, min_move_px: int, min_move_interval: float):
    events = []
    last_ts = time.time()
    last_move = {'x': None, 'y': None, 't': 0.0}
    running = {'on': True}  # truque para fechar listeners de dentro

    print("🎬 Gravando eventos (mouse + teclado).")
    print(f"   Para encerrar, pressione {STOP_KEY} (não será registrado).")
    print(f"   Salvando em: {Path(output_file).resolve()}")

    def push_event(ev_type: str, payload: dict):
        nonlocal last_ts
        now = time.time()
        delay = now - last_ts
        last_ts = now
        ev = {"type": ev_type, "delay": round(delay, 4)}
        ev.update(payload)
        events.append(ev)

    # Mouse
    def on_move(x, y):
        # filtro por distância mínima e intervalo mínimo
        now = time.time()
        if last_move['x'] is None:
            last_move['x'], last_move['y'], last_move['t'] = x, y, now
            push_event("mouse_move", {"x": int(x), "y": int(y)})
            return
        dx = abs(x - last_move['x'])
        dy = abs(y - last_move['y'])
        dt = now - last_move['t']
        if (dx + dy) >= min_move_px and dt >= min_move_interval:
            last_move['x'], last_move['y'], last_move['t'] = x, y, now
            push_event("mouse_move", {"x": int(x), "y": int(y)})

    def on_click(x, y, button, pressed):
        if not running['on']:
            return False
        # registra press (down) e release (up)
        b = norm_button(button)
        push_event("mouse_down" if pressed else "mouse_up",
                   {"x": int(x), "y": int(y), "button": b})

    def on_scroll(x, y, dx, dy):
        # dy: positivo = scroll up (para pyautogui.scroll positivo é para cima)
        push_event("mouse_scroll", {"x": int(x), "y": int(y), "dx": int(dx), "dy": int(dy)})

    # Teclado
    def on_press(k):
        if k == STOP_KEY:
            print("⏹ Parando gravação...")
            running['on'] = False
            return False  # encerra teclado
        nk = norm_key(k)
        if nk:
            push_event("key_down", {"key": nk})

    def on_release(k):
        if k == STOP_KEY:
            return False
        nk = norm_key(k)
        if nk:
            push_event("key_up", {"key": nk})

    ml = mouse.Listener(on_move=on_move, on_click=on_click, on_scroll=on_scroll)
    kl = keyboard.Listener(on_press=on_press, on_release=on_release)

    ml.start()
    kl.start()

    # Espera até teclado parar (F12)
    kl.join()
    # Para mouse também
    ml.stop()
    ml.join()

    # Ajusta delay do primeiro evento para 0
    if events:
        events[0]['delay'] = 0.0

    Path(output_file).parent.mkdir(parents=True, exist_ok=True)
    Path(output_file).write_text(json.dumps({"created_at": time.time(), "events": events}, indent=2), encoding="utf-8")
    print(f"✅ Gravação salva ({len(events)} eventos).")

# -------------- Reprodução --------------
def play(input_file: str, speed: float = 1.0, safe_wait: float = 3.0):
    path = Path(input_file)
    data = json.loads(path.read_text(encoding="utf-8"))
    events = data.get("events", [])
    if not events:
        print("⚠️ Nenhum evento para reproduzir.")
        return

    print(f"▶️ Reproduzindo {len(events)} eventos (speed={speed}).")
    if safe_wait and safe_wait > 0:
        print(f"⏳ Iniciando em {safe_wait:.0f}s... (mover mouse para o canto superior esquerdo aborta)")
        time.sleep(safe_wait)

    for i, ev in enumerate(events, 1):
        # respeita o delay (ajustado pela velocidade)
        time.sleep(max(0.0, ev.get('delay', 0.0)) / max(0.01, speed))

        t = ev.get("type")
        try:
            if t == "mouse_move":
                pyautogui.moveTo(ev["x"], ev["y"])
            elif t == "mouse_down":
                pyautogui.mouseDown(x=ev["x"], y=ev["y"], button=ev.get("button", "left"))
            elif t == "mouse_up":
                pyautogui.mouseUp(x=ev["x"], y=ev["y"], button=ev.get("button", "left"))
            elif t == "mouse_scroll":
                # vertical
                dy = int(ev.get("dy", 0))
                if dy:
                    pyautogui.scroll(dy)
                # horizontal (se disponível na sua versão do pyautogui)
                dx = int(ev.get("dx", 0))
                if dx:
                    try:
                        pyautogui.hscroll(dx)
                    except Exception:
                        pass
            elif t == "key_down":
                pyautogui.keyDown(ev["key"])
            elif t == "key_up":
                pyautogui.keyUp(ev["key"])
            else:
                # desconhecido: ignora
                pass
            print(f"[{i}/{len(events)}] {t}")
        except pyautogui.FailSafeException:
            print("⛔ Fail-safe ativado (mouse em 0,0). Interrompendo reprodução.")
            break

# -------------- Utilitários --------------
def show(input_file: str, limit: int | None = None):
    path = Path(input_file)
    data = json.loads(path.read_text(encoding="utf-8"))
    events = data.get("events", [])
    print(f"📄 {path} — {len(events)} eventos")
    for i, ev in enumerate(events[:limit] if limit else events, 1):
        print(f"{i:03d}: {ev}")

# -------------- CLI --------------
def main():
    p = argparse.ArgumentParser(description="Gravar e reproduzir mouse+teclado (PyAutoGUI + pynput)")
    mx = p.add_mutually_exclusive_group(required=True)
    mx.add_argument("--record", metavar="ARQ_JSON", help="Gravar e salvar em ARQ_JSON")
    mx.add_argument("--play", metavar="ARQ_JSON", help="Reproduzir a partir de ARQ_JSON")
    mx.add_argument("--show", metavar="ARQ_JSON", help="Listar eventos de ARQ_JSON")

    p.add_argument("--speed", type=float, default=1.0, help="Velocidade do play (1.0 = normal, 2.0 = 2x)")
    p.add_argument("--safe-wait", type=float, default=3.0, help="Espera inicial (s) antes do play")
    p.add_argument("--min-move", type=int, default=DEFAULT_MIN_MOVE_PX, help="Mín. px p/ registrar movimento")
    p.add_argument("--min-interval", type=float, default=DEFAULT_MIN_MOVE_INTERVAL, help="Mín. s entre movimentos")

    args = p.parse_args()

    if args.record:
        record(args.record, args.min_move, args.min_interval)
    elif args.play:
        play(args.play, speed=args.speed, safe_wait=args.safe_wait)
    elif args.show:
        show(args.show)

if __name__ == "__main__":
    main()