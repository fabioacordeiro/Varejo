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

STOP_KEY = keyboard.Key.esc
events = []
last_time = None
recording = True

def record(output_file):
    global last_time, recording

    def on_click(x, y, button, pressed):
        global last_time, recording
        if not recording:
            return False
        if pressed is False:  # registra apenas quando solta o botão
            now = time.time()
            delay = 0 if last_time is None else round(now - last_time, 4)
            last_time = now
            events.append({"x": x, "y": y, "button": str(button), "delay": delay})
            print(f"⏺ Clique registrado: {events[-1]}")

    def on_key(key):
        global recording
        if key == STOP_KEY:
            print("⏹ Gravação encerrada!")
            recording = False
            return False

    print("🎬 Gravando cliques... Pressione ESC para parar.")
    with mouse.Listener(on_click=on_click) as ml, keyboard.Listener(on_press=on_key) as kl:
        ml.join()
        kl.join()

    # garante delay zero no primeiro
    if events:
        events[0]["delay"] = 0.0

    Path(output_file).write_text(json.dumps(events, indent=2), encoding="utf-8")
    print(f"✅ Gravado com sucesso em: {output_file}")

def play(input_file):
    data = json.loads(Path(input_file).read_text(encoding="utf-8"))
    print(f"▶️ Reproduzindo {len(data)} cliques... (não mexa no mouse agora)")
    time.sleep(2)

    for ev in data:
        time.sleep(ev["delay"])
        pyautogui.click(x=ev["x"], y=ev["y"])
        print(f"🖱 Clique em ({ev['x']}, {ev['y']})")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--record", help="Gravar cliques em arquivo JSON")
    parser.add_argument("--play", help="Reproduzir cliques de um arquivo JSON")
    args = parser.parse_args()

    if args.record:
        record(args.record)
    elif args.play:
        play(args.play)
    else:
        print("Use --record ou --play")

if __name__ == "__main__":
    main()