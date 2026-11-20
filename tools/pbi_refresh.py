"""pbi_refresh.py - 열린 Power BI Desktop 에서 [홈 > 새로 고침] 을 눌러 데이터를 모두 불러온다 (교안 제작용)."""
import sys, time
sys.path.insert(0, __file__.rsplit("\\", 1)[0] if "\\" in __file__ else "tools")
import pbi_ui


def refresh(win=None):
    w = win or pbi_ui.main_window()
    w.set_focus()
    wr = w.rectangle()
    # 리본의 [새로 고침] 분할 단추 윗부분 (높이 37 짜리, 리본 줄 = 창 위에서 70~100 픽셀)
    cands = [b for b in w.descendants(control_type="Button")
             if b.window_text() == "새로 고침" and b.is_visible()
             and 60 < b.rectangle().top - wr.top < 100 and b.rectangle().height() < 50]
    cands[0].click_input()
    return w


if __name__ == "__main__":
    refresh()
    print("refresh clicked")
