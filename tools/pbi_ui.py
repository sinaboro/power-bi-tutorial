"""pbi_ui.py - Power BI Desktop 창을 다루는 도우미 (교안 화면 캡처·자동 조작용)

    python tools/pbi_ui.py shot 저장할파일.png      # 지금 Power BI Desktop 창을 캡처
    python tools/pbi_ui.py tree                     # 창 안의 UI 요소 목록 (자동 조작할 버튼 이름 찾기)

교안을 만들 때만 쓰는 도구라 학생은 실행하지 않아도 된다.
pywinauto 로 창을 찾고, PIL ImageGrab 으로 창 영역을 캡처한다.
"""
import sys
import time

from PIL import ImageGrab
from pywinauto import Desktop


def main_window(title_part=None, timeout=60):
    """PBIDesktop.exe 의 가장 큰 창을 찾는다."""
    end = time.time() + timeout
    while time.time() < end:
        wins = [w for w in Desktop(backend="uia").windows() if w.element_info.process_id and
                _proc_name(w.element_info.process_id) == "PBIDesktop.exe"]
        if title_part:
            wins = [w for w in wins if title_part in w.window_text()]
        if wins:
            return max(wins, key=lambda w: w.rectangle().width() * w.rectangle().height())
        time.sleep(1)
    raise RuntimeError("Power BI Desktop 창을 찾지 못했습니다")


_names = {}


def _proc_name(pid):
    if pid not in _names:
        import psutil
        try:
            _names[pid] = psutil.Process(pid).name()
        except Exception:
            _names[pid] = ""
    return _names[pid]


def shot(path, win=None, crop=None):
    """창을 앞으로 가져와 캡처한다. crop=(왼, 위, 오른, 아래) 창 기준 좌표로 일부만 자를 수 있다."""
    win = win or main_window()
    try:
        win.set_focus()
    except Exception:
        pass
    time.sleep(0.8)
    r = win.rectangle()
    img = ImageGrab.grab(bbox=(r.left, r.top, r.right, r.bottom), all_screens=True)
    if crop:
        img = img.crop(crop)
    img.save(path)
    return path


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "shot":
        print(shot(sys.argv[2]))
    elif cmd == "tree":
        w = main_window()
        print(w.window_text(), w.rectangle())
        w.print_control_identifiers(depth=int(sys.argv[2]) if len(sys.argv) > 2 else 4)
