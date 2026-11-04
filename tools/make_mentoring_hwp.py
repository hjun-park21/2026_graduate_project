# -*- coding: utf-8 -*-
"""3차 멘토링 자료(2차 자문 회신)를 한글 문서로 직접 짜고, 그 문서를 PDF 로 저장합니다. 담당 D.

    python -m tools.make_mentoring_hwp                 # docs/mentoring-3.hwp + .pdf
    python -m tools.make_mentoring_hwp --preview DIR   # 쪽마다 그림도 떠서 DIR 에 저장

**한글 → PDF 순서입니다.** 처음에는 마크다운 → HTML → 한글로 옮겼는데(`md_to_hwp`),
한글의 HTML 가져오기가 문단 간격·표 모양을 다 잃어 "표 밑에 문단이 다 붙은" 문서가
됐습니다. 그래서 이 도구는 HTML 을 거치지 않고 한컴오피스 자동화(COM)로 글자 모양,
문단 간격, 표를 하나씩 직접 만듭니다. PDF 는 그 한글 문서를 **한글이** 저장한 것이라
두 파일의 모양이 같습니다.

양식은 팀이 쓴 `다찾아드림_3차_멘토링.pdf`(노션 내보내기) 를 따랐습니다 — 큰 제목,
"1. / ①" 두 단계 제목, 머리행만 옅게 칠한 회색 선 표, 왼쪽 아래 문서 이름과 오른쪽
아래 쪽 번호.

**수치는 전부 원자료에서 왔습니다.** 아래 CONTENT 의 숫자를 고칠 때는 출처를 확인하고
`docs/PROGRESS.md` 도 같이 고치세요.
- 신호 조합 15개      `eval/ablation.json`   (2026-10-21, 48문항)
- 임계값 곡선         `eval/threshold_20261026.json`
- 병합 실험           `eval/merge_experiment.json` · `PROGRESS.md` 5-5장
- 처리 시간           `python -m tools.cost_report`
- 사용자 평가         `eval/user_study_raw.csv` · `PROGRESS.md` 5-7장
- 신뢰구간·부호검정   문항별 기록을 짝지어 부트스트랩 10,000회 (`PROGRESS.md` 5-2장)

⚠ 한글 API 의 속성 이름에 **오타가 있습니다** — 왼쪽 테두리 색은 `BorderCorlorLeft` 입니다.
  `BorderColorLeft` 로 쓰면 오류가 나고, 고치기 전에는 표 왼쪽 선만 검게 나왔습니다.
⚠ 표 열 너비(`ColWidth`)에는 셀 안쪽 여백이 **더해집니다.** 여백을 빼고 넣어야 표가
  본문 폭을 넘지 않습니다.
⚠ 끝나도 `Hwp.exe` 가 남을 수 있습니다(`md_to_hwp` 와 같음). 사용자가 한글로 다른
  문서를 열어 두었을 수 있어 강제로 죽이지 않습니다.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from server.config import ROOT

OUT = ROOT / "docs" / "mentoring-3.hwp"

FONT = "Noto Sans KR"
INK = (55, 53, 47)          # 본문 글자 (노션 기본 글자색)
MUTED = (120, 119, 116)     # 주석
FOOT = (155, 154, 151)      # 꼬리말
LINE = (222, 221, 218)      # 표 선
HEAD_FILL = (247, 246, 243) # 표 머리행

PAGE_W, MARGIN_LR = 210, 20
TEXT_W = PAGE_W - 2 * MARGIN_LR   # 본문 폭 170mm
CELL_PAD_LR = 2.5                 # 셀 안쪽 좌우 여백(mm)

# ---------------------------------------------------------------------------
# 내용. ("종류", ...) 를 순서대로 씁니다. 문단 안의 **굵게** 는 굵은 글씨가 됩니다.
#   표: ("table", 머리행, 행들, {"widths": [...], "bold": {행번호}, "keep": 앞문단과 붙이기})
# ---------------------------------------------------------------------------
CONTENT = [
    ("title", "3차 멘토링"),

    ("h1", "1. 보완 개요"),
    ("p", "2차 멘토링에서 제시된 의견을 바탕으로 장면 설명을 보강한 뒤 **공식 검색 성능을 재평가**하고, "
          "멀티모달 신호별 기여도와 **비용 대비 효과**를 분석하였다. 또한 근거 부족 질문 대응, 장면 병합 방식, "
          "처리 효율 등을 실험적으로 검토하고, 실제 사용자 평가를 통해 시스템의 활용 가능성을 검증하였다."),
    ("p", "평가에는 **외부 인원이 작성한 질문 48개**(재료·분량 19, 순서·타이밍 16, 시각적 상태 13)와 "
          "사람이 직접 판정한 정답 시간 구간을 사용하였으며, 대상은 저장된 영상 18개·장면 카드 203개이다."),

    ("h1", "2. 주요 자문 의견 및 보완 결과"),

    ("h2", "① 공식 기준선 확정 및 신호별 기여도 분석"),
    ("p", "2차 멘토링 당시 제시한 0.806은 장면 설명이 55%만 생성된 상태에서, 설명이 없는 영상 5개를 제외하고 "
          "측정한 값이었다. 저장된 프레임으로 장면 설명을 **81%까지 보강**한 뒤(구조상 상한 82%) 전체 영상 18개를 "
          "대상으로 재측정하였고, 음성·자막·장면 설명·이미지의 **15개 신호 조합을 전수 평가**하였다."),
    ("table", ["평가 조건", "nDCG@5", "Recall@5", "mAP", "1위적중"], [
        ["음성 단독 (기준선)", "0.717", "0.750", "0.610", "0.688"],
        ["전체 신호 결합", "0.745", "0.802", "0.621", "0.688"],
        ["장면 설명 제외 (최종 구성)", "0.772", "0.833", "0.660", "0.750"],
        ["자막 제외", "0.765", "0.781", "0.630", "0.708"],
        ["이미지 제외", "0.734", "0.781", "0.612", "0.646"],
        ["음성 제외", "0.697", "0.771", "0.561", "0.604"],
        ["이미지 단독", "0.470", "0.594", "0.368", "0.292"],
        ["무작위 선택 (참고)", "0.443", "–", "–", "–"],
    ], {"widths": [58, 28, 28, 28, 28], "bold": {2}}),
    ("note", "※ \"~ 제외\"는 전체 신호 결합에서 해당 신호 하나를 뺀 조건이다. 무작위 선택은 같은 영상 안에서 "
             "장면 카드 5개를 임의로 고른 값(2,000회 평균)이다."),
    ("p", "평가 결과, 전체 신호 결합 대비 **장면 설명 제외 조건에서 가장 높은 검색 성능**(nDCG@5 0.772)이 "
          "관찰되었으며, **네 가지 평가 지표 모두에서 동일한 개선 방향**을 보였다. 반면 음성 제외 시 성능 감소가 "
          "가장 크게 나타났고, 이미지 단독 검색(0.470)은 무작위 선택의 95% 범위(0.366~0.523) 안에 있어 "
          "영상 안의 특정 순간을 찾는 데에는 단독으로 기여하지 못함을 확인하였다."),
    ("p", "질문 유형별로 보면, 장면 설명은 **재료·분량 질문에서 성능을 통계적으로 유의하게 떨어뜨렸다**"
          "(제외 시 +0.084, 95% 신뢰구간 +0.037~+0.139). 분량 정보는 화면 자막에 글자로 표기되는데, "
          "생성된 설명이 이를 희석한 것으로 판단된다."),

    ("h2", "② 비용 대비 효과 재검토 및 최종 구성 결정"),
    ("p", "장면 설명은 시스템에서 **유일하게 외부 API(Gemini)를 호출하는 단계**로, 서버 처리 시간의 대부분을 "
          "차지하는 프레임 분석 단계에 포함되며 무료 API 한도상 하루 처리 가능한 영상 수를 20개로 제한한다. "
          "성능 측면에서도 ①과 같이 제외하는 편이 네 지표 모두에서 높았다. 이에 따라 **장면 설명을 제외한 "
          "음성·자막·이미지 결합 구성을 최종 구성으로 채택**하였다."),
    ("table", ["지표", "음성 단독", "최종 구성", "개선폭", "95% 신뢰구간"], [
        ["nDCG@5", "0.717", "0.772", "+0.055", "+0.005 ~ +0.109"],
        ["Recall@5", "0.750", "0.833", "+0.083", "+0.031 ~ +0.135"],
        ["mAP", "0.610", "0.660", "+0.050", "−0.003 ~ +0.110"],
        ["1위적중", "68.8%", "75.0%", "+6.2%p", "−4.2 ~ +16.7%p"],
    ], {"widths": [30, 30, 30, 30, 50], "bold": set()}),
    ("p", "음성 단독 대비 최종 구성은 **nDCG@5와 Recall@5에서 신뢰구간이 0을 벗어나 개선이 확인**되었고, "
          "mAP와 1위적중은 같은 방향이나 현재 표본에서는 확정되지 않았다. 이 구성에서는 색인 단계의 외부 API "
          "호출이 없어져 하루 처리량 제약도 함께 해소된다."),
    ("p", "다만 두 가지 한계가 있다. 첫째, 장면 설명이 기여할 것으로 기대되는 **시각적 상태 질문**에서는 설명을 "
          "제외하면 nDCG@5가 0.046 낮아져 손해 방향이었으나, 해당 유형 질문이 13개뿐이어서 신뢰구간"
          "(−0.117~+0.013)이 0을 포함한다. "
          "둘째, 최종 구성은 15개 조합 중 **같은 평가셋에서 가장 높은 조합을 고른 것**이므로 수치가 다소 "
          "낙관적일 수 있다. 시각적 상태 질문을 보강한 새 질문으로 재확인할 예정이다."),

    ("h2", "③ 답이 없는 질문에 대한 대응"),
    ("p", "관련 장면이 없는 질문에도 검색 결과가 제공되던 문제를 개선하기 위해 **정규화 이전 유사도 기반의 "
          "근거 부족 경고 기능**을 구현하였다. 점수를 정규화하면 1위가 항상 만점이 되어 확신도를 표현할 수 "
          "없으므로, 정규화 전 코사인 유사도를 별도로 보존해 판정에 사용하였다. 운영 임계값은 아래 곡선에서 "
          "**데이터로 선택**하였다."),
    ("table", ["임계값", "분야가 다른 질문\n(10건 중 경고)", "정답 장면이 없는 질문\n(12건 중 경고)",
               "정답이 있는 질문\n(48건 중 불필요한 경고)"], [
        ["0.48", "8", "2", "2"],
        ["0.50", "9", "4", "4"],
        ["0.52 (채택)", "10", "7", "5"],
        ["0.54", "10", "9", "8"],
        ["0.58", "10", "12", "21"],
    ], {"widths": [30, 44, 48, 48], "bold": {2}}),
    ("p", "임계값 0.52 적용 결과, **분야가 다른 질문 10건 모두와 정답 장면이 없는 질문 12건 중 7건**에서 경고가 "
          "발생하였다. 다만 정답이 있는 질문에서도 **5건의 불필요한 경고**가 발생하여, 결과를 차단하지 않고 "
          "**경고와 함께 후보 장면을 제공하는 방식**을 채택하였다. 차단하면 실제로 존재하는 장면을 사용자가 "
          "볼 수 없게 되지만, 경고에 그치면 피해가 안내 문구 한 줄로 끝나기 때문이다."),

    ("h2", "④ 순서·타이밍 질문의 검색 정확도 개선"),
    ("p", "4초 고정 분할의 한계를 개선하기 위해 **인접 장면 병합 및 시점 보정 기능을 구현하고 성능을 비교**"
          "하였다."),
    ("table", ["분할 방식", "카드 길이", "정답 포함 장면 검색률", "평균 시점 오차"], [
        ["4초 고정 분할 (유지)", "4.0초", "69% (65/94)", "2.8초"],
        ["인접 장면 병합", "5.4초", "78% (73/94)", "3.5초"],
    ], {"widths": [50, 30, 50, 40], "bold": {0}}),
    ("p", "병합 시 정답 구간을 포함하는 장면 검색률은 69%에서 78%로 증가했으나, 카드가 길어진 만큼 이동 지점이 "
          "뒤로 밀려 평균 시점 오차가 2.8초에서 3.5초로 증가하면서 최종 검색 성능은 저하되었다. 시점 보정과 "
          "이상적 조건의 추가 실험에서도 실질적인 개선을 확인하지 못해, 현재는 **시점 이동의 정확성을 고려하여 "
          "4초 고정 분할 방식을 유지**하기로 결정하였다."),

    ("h2", "⑤ 영상 처리 시간 및 비용 분석"),
    ("p", "영상 18개를 대상으로 저장 과정의 단계별 소요 시간을 실측하였다."),
    ("table", ["구분", "평균 소요 시간", "비중"], [
        ["영상 캡처 (실시간 재생)", "47.2초", "전체의 30%"],
        ["서버 처리", "111.7초", "전체의 70%"],
        ["   음성 인식", "13.7초", "서버의 12%"],
        ["   프레임 분석 (자막 인식 + 장면 설명)", "90.8초", "서버의 81%"],
        ["   벡터 생성 및 기타", "7.2초", "서버의 7%"],
        ["합계 (사용자 대기 시간)", "158.9초", "100%"],
    ], {"widths": [90, 40, 40], "bold": {5}}),
    ("p", "영상당 평균 **158.9초**가 소요되었으며, **서버 처리 시간의 81%가 프레임 분석 단계에 집중**되어 있음을 "
          "확인하였다. 외부 API 호출을 영상당 4회에서 1회로 축소하였으나, 무료 API의 일일 호출 제한으로 하루 20개 "
          "영상이 처리량 상한으로 남아 있었다. ②의 최종 구성에서는 장면 설명이 빠져 색인 단계의 외부 호출이 "
          "없어지므로, 남은 과제는 **프레임 분석(자막 인식) 단계의 최적화**이다."),

    ("h2", "⑥ 실제 사용자 평가를 통한 활용 가능성 검증"),
    ("p", "참가자 10명을 대상으로 총 60개 과제를 수행하여 YouTube 직접 탐색과 제안 시스템의 탐색 효율을 "
          "비교하였다."),
    ("table", ["항목", "YouTube 직접 탐색", "제안 시스템"], [
        ["평균 탐색 시간", "27.3초", "9.3초"],
        ["중앙값", "27.3초", "8.5초"],
        ["평균 시도 횟수", "1.56회", "1.10회"],
        ["과제 성공률 (사전 기준)", "93% (28/30)", "83% (25/30)"],
    ], {"widths": [60, 55, 55], "bold": {0}}),
    ("p", "평가 결과, **평균 탐색 시간이 27.3초에서 9.3초로 약 66% 단축**되었으며, 참가자 10명 전원에서 시간 "
          "감소가 관찰되었다(부호검정 **p = 0.002**). 과제 성공률은 두 방식 모두 90% 안팎으로 높아 차이를 가리기 "
          "어려웠다. 두 방식의 실패 7건 중 6건은 정답 구간 안에 도달했으나 사전 기준(정답 시작 −5초~+2초)을 "
          "벗어난 경우로, 정답 구간이 긴 과제 2개에서 발생하였다. 따라서 이번 평가에서 확인된 효과는 "
          "**성공률 향상이 아니라 탐색 시간 단축**이다."),

    ("h1", "3. 종합 결과"),
    ("table", ["2차 자문 의견", "보완 결과"], [
        ["장면 설명 생성 후 공식 기준선 확정", "설명 81% 보강 후 15개 신호 조합 전수 평가 (①)"],
        ["비용 대비 효과 재검토", "장면 설명 제외 구성 채택, 음성 대비 nDCG@5 0.717 → 0.772 (②)"],
        ["답이 없는 질문 대응", "임계값 0.52 근거 부족 경고, 운영점을 곡선으로 선택 (③)"],
        ["순서·타이밍 개선 (조리 단계 병합)", "구현·비교 후 미채택, 시점 오차 증가 (④)"],
        ["처리 시간·비용 측정", "영상당 158.9초, 프레임 분석이 서버의 81% (⑤)"],
        ["사용자 평가로 서비스 가능성 제시", "탐색 시간 66% 단축, 10명 전원 개선, p = 0.002 (⑥)"],
    ], {"widths": [62, 108], "bold": set()}),
    ("p", "2차 멘토링 이후 검색 성능 평가 체계를 정비하고, 멀티모달 신호별 기여도와 처리 비용, 장면 분할 방식 및 "
          "근거 부족 질문 대응을 실험적으로 검증하였다. 특히 **각 기술의 실제 효과를 정량적으로 비교하고, 결과에 "
          "따라 개선안의 채택 여부를 결정**함으로써 시스템 설계의 타당성을 검토하였다. 또한 실제 사용자 평가에서 "
          "평균 탐색 시간을 약 66% 단축하여, **저장된 숏폼 콘텐츠의 정보 접근성과 활용성을 높일 수 있는 가능성**을 "
          "확인하였다."),

    ("h1", "4. 향후 과제"),
    ("table", ["과제", "목적"], [
        ["시각적 상태 질문 보강(13개 → 30개) 후 새 질문으로 재측정", "최종 구성의 성능을 독립적으로 재확인 (②)"],
        ["근거 부족 경고·처리 시간·사용자 평가를 최종 구성으로 재측정", "③⑤⑥은 장면 설명을 포함한 구성에서 측정됨"],
        ["정답 장면이 없는 질문 보강", "경고 임계값 0.52 정밀화 (③)"],
        ["프레임 분석(자막 인식) 단계 최적화", "서버 처리 시간 단축 (⑤)"],
    ], {"widths": [98, 72], "bold": set()}),

    # 역할 분담표 — docs/plan-and-roles.md 2장과 같은 내용. 담당자 칸은 일부러 비워 둡니다.
    ("h1", "5. 역할 분담"),
    ("table", ["구분", "담당 업무", "주요 산출물", "담당자"], [
        ["A", "**시스템 설계 · 서버 백엔드**\n파이프라인 전 구간 설계, FastAPI 서버, 작업 큐, SQLite·Qdrant 연동, 사용자 격리",
         "서버 전 구간, 비교 실험 러너", ""],
        ["B", "**분석 파이프라인 안정화**\n단계별 상태 기록, 실패 등급 구분, 할당량 오류 처리, 로컬 모델 동시 실행 제어, 자동 테스트",
         "실패를 드러내는 구조, 테스트 66건", ""],
        ["C", "**멀티모달 모델 · 검색 품질**\nSigLIP2 파인튜닝 실험, 질의 전처리, 하이브리드 검색 가중치, 조용한 실패 버그 추적",
         "실험 기록, 검색 로직 개선", ""],
        ["D", "**평가 체계 · 자문 대응**\n평가셋 설계, 측정 도구, 신호별 기여도·병합·거절 임계값 실험, 처리 시간·비용 실측, 사용자 평가 설계",
         "평가 스크립트 6종, 측정 원자료, 통합 문서", ""],
        ["E", "**크롬 확장 UI/UX**\n저장 흐름, 검색 화면, 사진 레시피 뷰, 재료 구매 링크, 제목·채널 추출",
         "확장 전 구간", ""],
        ["공통", "**평가셋 라벨링**\n담당 영상의 정답 구간을 사람이 직접 판정 (영상 4개씩 분담)", "라벨 113줄", ""],
        ["공통", "**사용자 평가 진행**\n참가자 섭외, 진행, 기록 (진행자 4명이 10명 담당)", "60과제 원자료", ""],
        ["공통", "**질문 수집**\n팀 외부 지인에게 질문 의뢰 — 개발팀이 만들면 \"잘 푸는 질문\"으로 기울어짐",
         "질문 64개", ""],
    ], {"widths": [14, 92, 44, 20], "bold": set()}),
]


# ---------------------------------------------------------------------------
# 한글 자동화
# ---------------------------------------------------------------------------
class Hwp:
    def __init__(self) -> None:
        import win32com.client
        self.hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
        try:
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        self.A, self.P = self.hwp.HAction, self.hwp.HParameterSet

    def mm(self, v: float) -> int:
        return self.hwp.MiliToHwpUnit(v)

    def rgb(self, c: tuple[int, int, int]) -> int:
        return self.hwp.RGBColor(*c)

    # --- 쪽 -----------------------------------------------------------------
    def page(self) -> None:
        sd = self.P.HSecDef
        self.A.GetDefault("PageSetup", sd.HSet)
        pd = sd.PageDef
        pd.LeftMargin = pd.RightMargin = self.mm(MARGIN_LR)
        pd.TopMargin, pd.BottomMargin = self.mm(20), self.mm(14)
        pd.HeaderLen, pd.FooterLen = self.mm(0), self.mm(10)
        sd.HSet.SetItem("ApplyClass", 24)
        sd.HSet.SetItem("ApplyTo", 3)
        self.A.Execute("PageSetup", sd.HSet)

    def footer(self, text: str) -> None:
        hf = self.P.HHeaderFooter
        self.A.GetDefault("HeaderFooter", hf.HSet)
        hf.HSet.SetItem("HeaderFooterStyle", 0)
        hf.HSet.SetItem("HeaderFooterCtrlType", 1)   # 1 = 꼬리말
        self.A.Execute("HeaderFooter", hf.HSet)
        self.para(line=130)
        self.char(8, color=FOOT)
        self.text(text)
        self.A.Run("CloseEx")
        pn = self.P.HPageNumPos
        self.A.GetDefault("PageNumPos", pn.HSet)
        pn.DrawPos = self.hwp.PageNumPosition("BottomRight")
        pn.NumberFormat = self.hwp.NumberFormat("Digit")
        pn.SideChar = 0     # 기본값이면 "- 1 -" 처럼 줄표가 붙습니다
        self.A.Execute("PageNumPos", pn.HSet)

    # --- 글자·문단 ------------------------------------------------------------
    def char(self, size: float, bold: bool = False, color=INK) -> None:
        cs = self.P.HCharShape
        self.A.GetDefault("CharShape", cs.HSet)
        for k in ("Hangul", "Latin", "Hanja", "Japanese", "Other", "Symbol", "User"):
            cs.HSet.SetItem("FaceName" + k, FONT)
            cs.HSet.SetItem("FontType" + k, 1)
            cs.HSet.SetItem("Spacing" + k, -2)
        cs.Height = int(size * 100)
        cs.Bold = 1 if bold else 0
        cs.TextColor = self.rgb(color)
        self.A.Execute("CharShape", cs.HSet)

    def para(self, line: int = 165, before: float = 0, after: float = 0,
             align: int = 1, keep: bool = False) -> None:
        """align: 0 양쪽 · 1 왼쪽 · 2 오른쪽 · 3 가운데. before/after 는 pt."""
        ps = self.P.HParaShape
        self.A.GetDefault("ParagraphShape", ps.HSet)
        ps.LineSpacingType = 0
        ps.LineSpacing = line
        ps.PrevSpacing = int(before * 200)
        ps.NextSpacing = int(after * 200)
        ps.AlignType = align
        ps.KeepWithNext = 1 if keep else 0
        ps.LeftMargin = 0
        ps.Indentation = 0
        self.A.Execute("ParagraphShape", ps.HSet)

    def text(self, s: str) -> None:
        it = self.P.HInsertText
        self.A.GetDefault("InsertText", it.HSet)
        it.Text = s
        self.A.Execute("InsertText", it.HSet)

    def rich(self, s: str, size: float, color=INK, bold_all: bool = False) -> None:
        """`**굵게**` 표시를 굵은 글씨로 바꿔 넣습니다."""
        for i, part in enumerate(re.split(r"\*\*", s)):
            if part:
                self.char(size, bold=bold_all or i % 2 == 1, color=color)
                self.text(part)

    def br(self) -> None:
        self.A.Run("BreakPara")

    # --- 표 -------------------------------------------------------------------
    def table(self, head: list[str], rows: list[list[str]], opt: dict) -> None:
        widths = opt["widths"]
        assert abs(sum(widths) - TEXT_W) < 0.5, f"표 너비 합 {sum(widths)} ≠ 본문 폭 {TEXT_W}"
        bold_rows = opt.get("bold", set())
        n_rows, n_cols = len(rows) + 1, len(head)

        self.para(line=100, before=3, after=5)
        tc = self.P.HTableCreation
        self.A.GetDefault("TableCreate", tc.HSet)
        tc.Rows, tc.Cols = n_rows, n_cols
        tc.WidthType, tc.HeightType = 2, 0
        tc.WidthValue = self.mm(TEXT_W)
        tc.CreateItemArray("ColWidth", n_cols)
        for i, w in enumerate(widths):
            # ⚠ 셀 안쪽 여백이 열 너비에 더해지므로 미리 뺍니다.
            tc.ColWidth.SetItem(i, self.mm(w - 2 * CELL_PAD_LR))
        tp = tc.TableProperties
        tp.TreatAsChar = 1
        tp.CellMarginLeft = tp.CellMarginRight = self.mm(CELL_PAD_LR)
        tp.CellMarginTop = tp.CellMarginBottom = self.mm(1.4)
        self.A.Execute("TableCreate", tc.HSet)

        for r, row in enumerate([head] + rows):
            for c, value in enumerate(row):
                bold = r == 0 or (r - 1) in bold_rows
                for k, line in enumerate(value.split("\n")):
                    if k:
                        self.br()
                    self.para(line=140)
                    self.char(9.5, bold=bold, color=INK)   # 빈 칸에도 글자 모양은 맞춰 둡니다
                    self.rich(line, 9.5, bold_all=bold)
                if not (r == n_rows - 1 and c == n_cols - 1):
                    self.A.Run("TableRightCell")

        # 전체 선택 → 옅은 회색 선
        self.A.Run("TableCellBlock")
        self.A.Run("TableCellBlockExtend")
        self.A.Run("TableCellBlockExtend")
        cb = self.P.HCellBorderFill
        self.A.GetDefault("CellBorderFill", cb.HSet)
        g = self.rgb(LINE)
        cb.BorderTypeLeft = cb.BorderTypeRight = cb.BorderTypeTop = cb.BorderTypeBottom = 1
        cb.BorderWidthLeft = cb.BorderWidthRight = cb.BorderWidthTop = cb.BorderWidthBottom = 0
        cb.BorderCorlorLeft = g     # ⚠ 한글 API 쪽 오타 그대로입니다
        cb.BorderColorRight = cb.BorderColorTop = cb.BorderColorBottom = g
        cb.TypeHorz = cb.TypeVert = 1
        cb.WidthHorz = cb.WidthVert = 0
        cb.ColorHorz = cb.ColorVert = g
        self.A.Execute("CellBorderFill", cb.HSet)
        self.A.Run("Cancel")

        # 머리행만 옅게 칠하기
        self.A.Run("TableColBegin")
        self.A.Run("TableColPageUp")
        self.A.Run("TableCellBlock")
        self.A.Run("TableCellBlockExtend")
        self.A.Run("TableColEnd")
        cb = self.P.HCellBorderFill
        self.A.GetDefault("CellBorderFill", cb.HSet)
        cb.FillAttr.type = self.hwp.BrushType("NullBrush|WinBrush")
        cb.FillAttr.WinBrushFaceColor = self.rgb(HEAD_FILL)
        cb.FillAttr.WinBrushHatchColor = self.rgb((0, 0, 0))
        cb.FillAttr.WinBrushFaceStyle = self.hwp.HatchStyle("None")
        cb.FillAttr.WindowsBrush = 1
        self.A.Execute("CellBorderFill", cb.HSet)
        self.A.Run("Cancel")

        # 표 밖으로 나와 새 문단
        self.hwp.MovePos(3)
        self.br()

    # --- 저장 -----------------------------------------------------------------
    def save(self, out: Path, preview: Path | None) -> int:
        out = out.resolve()
        out.parent.mkdir(parents=True, exist_ok=True)   # 한글은 없는 폴더에 저장하지 못합니다
        if not self.hwp.SaveAs(str(out), "HWP", ""):
            raise RuntimeError("한글 파일로 저장하지 못했습니다")
        pdf = out.with_suffix(".pdf")
        if not self.hwp.SaveAs(str(pdf), "PDF", ""):
            raise RuntimeError("PDF 로 저장하지 못했습니다")
        pages = self.hwp.PageCount
        if preview:
            preview.mkdir(parents=True, exist_ok=True)
            for i in range(pages):
                self.hwp.CreatePageImage(str(preview / f"page{i + 1}.bmp"), i, 100, 24, "bmp")
        return pages

    def close(self) -> None:
        try:
            self.hwp.Clear(1)
            self.hwp.Quit()
        except Exception:
            pass


def build(doc: Hwp) -> None:
    doc.page()
    doc.footer("3차 멘토링")
    for kind, *rest in CONTENT:
        if kind == "title":
            doc.para(line=130, before=0, after=10)
            doc.char(26, bold=True)
            doc.text(rest[0])
            doc.br()
        elif kind == "h1":
            doc.para(line=140, before=14, after=5, keep=True)
            doc.char(17, bold=True)
            doc.text(rest[0])
            doc.br()
        elif kind == "h2":
            doc.para(line=140, before=10, after=4, keep=True)
            doc.char(13, bold=True)
            doc.text(rest[0])
            doc.br()
        elif kind == "p":
            doc.para(line=170, before=0, after=6)
            doc.rich(rest[0], 10.5)
            doc.br()
        elif kind == "note":
            doc.para(line=150, before=0, after=6)
            doc.rich(rest[0], 8.8, color=MUTED)
            doc.br()
        elif kind == "table":
            doc.table(*rest)
        else:
            raise ValueError(kind)


def main() -> int:
    p = argparse.ArgumentParser(description="3차 멘토링 자료를 한글로 짜고 PDF 로 저장")
    p.add_argument("--out", default=str(OUT), help="한글 파일 경로 (PDF 는 같은 이름)")
    p.add_argument("--preview", default=None, help="쪽 그림을 저장할 폴더 (검토용)")
    args = p.parse_args()

    doc = Hwp()
    try:
        build(doc)
        pages = doc.save(Path(args.out), Path(args.preview) if args.preview else None)
    finally:
        doc.close()
    out = Path(args.out).resolve()
    for f in (out, out.with_suffix(".pdf")):
        if not f.is_file():
            print(f"만들어지지 않았습니다: {f}", file=sys.stderr)
            return 1
    print(f"만들었습니다: {out}  ({pages}쪽)")
    print(f"           {out.with_suffix('.pdf')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
