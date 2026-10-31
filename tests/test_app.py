from io import BytesIO
from pathlib import Path

import fitz
from docx import Document
from fastapi.testclient import TestClient
from pptx import Presentation
import pytest

import app as server

client = TestClient(server.app)


def test_home_and_health():
    r = client.get("/")
    assert r.status_code == 200
    assert "문제공방" in r.text
    assert client.get("/api/health").json()["status"] == "ok"


def test_parse_user_input_preserves_original():
    locked, wants, concepts = server.parse_user_input(
        "TBL은 뭐의 약자인가? → 답: Triple Bottom Line\n포터 가설\nTBL은 뭐의 약자인가? → 답: Triple Bottom Line",
        "유해물질저감기법\nMFCA\nMFCA",
    )
    assert locked == [{"question": "TBL은 뭐의 약자인가?", "answer": "Triple Bottom Line"}]
    assert wants == ["포터 가설"]
    assert concepts == ["유해물질저감기법", "MFCA"]


def test_extract_pdf_page_number():
    d = fitz.open()
    p = d.new_page()
    p.insert_text((50, 50), "Environmental management and sustainability")
    out = server.extract_segments("lecture.pdf", d.tobytes())
    assert len(out) == 1
    assert "p.1" in out[0]["label"]
    assert "Environmental" in out[0]["text"]


def test_extract_docx():
    d = Document()
    d.add_paragraph("Triple Bottom Line")
    f = BytesIO(); d.save(f)
    out = server.extract_segments("notes.docx", f.getvalue())
    assert "Triple Bottom Line" in out[0]["text"]


def test_extract_pptx():
    p = Presentation()
    slide = p.slides.add_slide(p.slide_layouts[1])
    slide.shapes.title.text = "Sustainability"
    f = BytesIO(); p.save(f)
    out = server.extract_segments("slides.pptx", f.getvalue())
    assert "Sustainability" in out[0]["text"]


def test_evidence_verification():
    lookup = {"S1": {"text": "기업은 경제적 성과와 환경적 성과를 균형 있게 달성한다."}}
    q = {"source_id": "S1", "evidence": "경제적 성과와 환경적 성과"}
    assert server.grounded(q, lookup)[0]
    assert not server.grounded({"source_id": "S1", "evidence": "전혀 다른 주장입니다"}, lookup)[0]


def test_generation_pinned_only_without_api_key(monkeypatch):
    server.visits.clear()
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    r = client.post("/api/generate", data={
        "past_text": "TBL은 뭐의 약자인가? → 답: Triple Bottom Line",
        "count": "5",
    })
    assert r.status_code == 200, r.text
    q = r.json()["questions"]
    assert q[0]["question"] == "TBL은 뭐의 약자인가?"
    assert q[0]["answer"] == "Triple Bottom Line"
    assert not q[0]["verified"]


def test_generation_with_mock_ai(monkeypatch):
    server.visits.clear()
    async def fake_ai(sources, pinned, wants, concepts, count, question_style, variants):
        return [{
            "question": "제품의 기능을 환경부하로 나눈 것은?",
            "answer": "환경경제효율성",
            "kind": "short", "options": [], "alternatives": [],
            "source_id": "S1", "evidence": "제품의 기능을 환경부하로 나눈다",
            "priority": "keyword"
        }]
    monkeypatch.setattr(server, "generate_ai", fake_ai)
    r = client.post("/api/generate", data={"count": "5", "concepts_text": "환경경제효율성"},
                    files={"files": ("lecture.txt", "제품의 기능을 환경부하로 나눈다. 환경경제효율성", "text/plain")})
    assert r.status_code == 200, r.text
    assert r.json()["questions"][0]["verified"]


def test_invalid_file_rejected():
    server.visits.clear()
    r = client.post("/api/generate", data={"count":"5"},
                    files={"files": ("malware.exe", b"123", "application/octet-stream")})
    assert r.status_code == 400


def test_source_round_robin():
    segments = [
        {"label": "a.pdf · p.1", "text": "aaa"},
        {"label": "a.pdf · p.2", "text": "bbb"},
        {"label": "b.pdf · p.1", "text": "ccc"},
    ]
    out, trunc = server.select_sources(segments)
    assert [s["label"] for s in out] == ["a.pdf · p.1", "b.pdf · p.1", "a.pdf · p.2"]
    assert not trunc
