# -*- coding: utf-8 -*-
"""마크다운 문서를 한글(.hwp) 로 뽑습니다. 담당 D.

    python -m tools.md_to_hwp docs/mentoring-3.md
    python -m tools.md_to_hwp docs/mentoring-3.md --out docs/mentoring-3.hwp

`.hwp` 는 공개 규격이 아니라 **파이썬으로 직접 쓸 수 없습니다.** 그래서 이 도구는
설치된 한컴오피스를 COM 으로 불러, `md_to_pdf` 가 만드는 것과 **같은 HTML** 을 열게 한 뒤
다른 이름으로 저장시킵니다. 한글이 깔려 있지 않은 컴퓨터에서는 동작하지 않습니다.

⚠ **보안 승인 창이 뜰 수 있습니다.** 한글은 스크립트가 파일을 열고 저장할 때 사용자에게
물어봅니다. 창이 뜨면 "허용" 을 누르면 되고, 그 전까지 이 명령은 기다립니다.
`FilePathCheckerModule` 이 등록돼 있으면 묻지 않고 지나갑니다(자동으로 시도합니다).

표 테두리·굵은 글씨·제목 단계는 HTML 쪽 CSS 를 그대로 따라갑니다. 한글에서 다시 손보려면
`md_to_pdf.CSS` 를 고치는 쪽이 빠릅니다 — 두 출력이 같은 모양을 유지합니다.

⚠ **변환이 끝나도 `Hwp.exe` 가 창 없이 남아 있을 때가 있습니다.** `Quit()` 을 부르지만
한글이 바로 죽지 않는 경우가 있어서입니다. 다음 변환에 지장은 없지만, 거슬리면 작업
관리자에서 끄면 됩니다. 사용자가 한글로 다른 문서를 열어 둔 채일 수 있으므로 이 도구는
`Hwp.exe` 를 강제로 죽이지 않습니다 — 저장 안 한 남의 문서가 날아갑니다.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from tools.md_to_pdf import CSS, render


def unwrap(body: str) -> str:
    """문단 안의 줄바꿈을 공백으로 바꿉니다.

    ⚠ 브라우저는 HTML 안의 줄바꿈을 공백으로 치지만 **한글은 그냥 지웁니다.** 그래서
    마크다운에서 문단을 적당한 길이로 접어 쓴 것이 `장면 병합방식` 처럼 붙어 버립니다.
    `<pre>` 안은 줄바꿈이 뜻을 가지므로 건드리지 않습니다.
    """
    out, last = [], 0
    for m in re.finditer(r"<pre\b.*?</pre>", body, flags=re.S):
        out.append(body[last:m.start()].replace("\n", " "))
        out.append(m.group(0))
        last = m.end()
    out.append(body[last:].replace("\n", " "))
    return "".join(out)


def write_html(body: str, path: Path) -> None:
    # ⚠ 한글은 `<meta charset="utf-8">` 를 **읽지 못합니다.** 그 한 줄로는 cp949 로
    #   잘못 읽어 본문이 통째로 깨지고(◆◆주題삯…), 덩달아 태그까지 밀려 `</td>` 같은
    #   것이 글자로 새어 나옵니다. 옛 `http-equiv` 형식은 알아봅니다.
    #   BOM(utf-8-sig)을 붙여도 인코딩은 살지만 맨 앞에 `&#65279;` 가 남습니다.
    path.write_text(
        "<!doctype html><html><head>"
        '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
        f"<style>{CSS}</style></head><body>{body}</body></html>",
        encoding="utf-8")


def verify(hwp, out: Path, probes: list[str]) -> None:
    """만든 문서를 텍스트로 되받아 글자가 성한지 봅니다.

    .hwp 는 열어 보기 전에는 깨졌는지 알 수 없어, 한글에게 같은 문서를 텍스트로도
    저장시켜 **본문에 있던 말이 그대로 있는지** 확인합니다. 인코딩이 틀어지면
    한 글자도 맞지 않으므로 이 검사에 바로 걸립니다.
    """
    txt = out.with_suffix(".verify.txt")
    try:
        hwp.SaveAs(str(txt.resolve()), "TEXT", "")
        raw = txt.read_bytes()
        for enc in ("utf-16", "cp949", "utf-8"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise RuntimeError("변환 결과를 읽을 수 없습니다")
        missing = [s for s in probes if s and s not in text]
        if missing:
            raise RuntimeError(f"본문이 깨졌습니다 — 찾지 못한 말: {missing[:3]}")
        if re.search(r"</(?:td|tr|p|table|strong)>", text):
            raise RuntimeError("HTML 태그가 글자로 새어 나왔습니다")
    finally:
        txt.unlink(missing_ok=True)


def to_hwp(html: Path, out: Path, probes: list[str]) -> None:
    # ⚠ 한글은 **절대 경로**만 받습니다. 상대 경로를 주면 조용히 실패합니다.
    import win32com.client

    hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
    try:
        # 보안 모듈이 등록돼 있으면 승인 창을 건너뜁니다. 없으면 그냥 넘어갑니다.
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        if not hwp.Open(str(html.resolve()), "HTML", ""):
            raise RuntimeError("한글이 HTML 을 열지 못했습니다")
        if not hwp.SaveAs(str(out.resolve()), "HWP", ""):
            raise RuntimeError("한글이 .hwp 로 저장하지 못했습니다")
        verify(hwp, out, probes)
    finally:
        try:
            hwp.Quit()
        except Exception:
            pass


def probes_from(src: Path) -> list[str]:
    """본문에서 검사에 쓸 한글 토막 3개를 뽑습니다 (처음·중간·끝)."""
    words = [w for w in re.findall(r"[가-힣]{4,}", src.read_text(encoding="utf-8"))]
    if not words:
        return []
    return [words[0], words[len(words) // 2], words[-1]]


def main() -> int:
    p = argparse.ArgumentParser(description="마크다운 → 한글(.hwp)")
    p.add_argument("files", nargs="+", help="변환할 .md 파일")
    p.add_argument("--out", default=None, help="출력 경로 (파일 하나일 때만)")
    args = p.parse_args()

    if args.out and len(args.files) > 1:
        print("--out 은 파일 하나일 때만 쓸 수 있습니다.", file=sys.stderr)
        return 2

    for name in args.files:
        src = Path(name)
        if not src.is_file():
            print(f"없는 파일: {src}", file=sys.stderr)
            return 1
        out = Path(args.out) if args.out else src.with_suffix(".hwp")
        out.parent.mkdir(parents=True, exist_ok=True)
        tmp = out.with_suffix(".tmp.html")
        write_html(unwrap(render(src)), tmp)
        try:
            to_hwp(tmp, out, probes_from(src))
        finally:
            tmp.unlink(missing_ok=True)
        out = out.resolve()
        if not out.is_file():
            print(f"한글 파일이 만들어지지 않았습니다: {out}", file=sys.stderr)
            return 1
        print(f"만들었습니다: {out}   ({src.name} · {out.stat().st_size // 1024}KB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
