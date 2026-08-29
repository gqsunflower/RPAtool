"""
疑似ローカルAIのコア: IntentEngine

外部LLM/クラウドAPIには一切接続しません。
ユーザーの自然文っぽい指示を、config/intents.json に登録された
キーワード・正規表現パターンとの一致度でスコアリングし、
最もスコアの高い intent (=実行すべきマクロ) を選び出します。

これは「AIっぽく振る舞うルールベースの意図分類器」であり、
本物の言語理解ではない点に留意してください(=疑似AI)。
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


@dataclass
class IntentMatch:
    intent_id: str
    macro: str
    score: float
    description: str


class IntentEngine:
    def __init__(self, intents_path: Path):
        self.intents_path = Path(intents_path)
        self._load()

    def _load(self) -> None:
        with open(self.intents_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.threshold: float = data.get("match_threshold", 1)
        self.intents: list[dict] = data.get("intents", [])

    def reload(self) -> None:
        """設定ファイルを再読込する(マクロや意図を追加した後に使用)。"""
        self._load()

    def _score(self, text: str, intent: dict) -> float:
        score = 0.0
        for kw in intent.get("keywords", []):
            if kw.lower() in text.lower():
                score += 1.0
        for pat in intent.get("patterns", []):
            try:
                if re.search(pat, text, flags=re.IGNORECASE):
                    score += 2.0
            except re.error:
                continue
        return score

    def classify(self, text: str) -> Optional[IntentMatch]:
        best: Optional[IntentMatch] = None
        best_score = 0.0
        for intent in self.intents:
            s = self._score(text, intent)
            if s > best_score:
                best_score = s
                best = IntentMatch(
                    intent_id=intent["id"],
                    macro=intent["macro"],
                    score=s,
                    description=intent.get("description", ""),
                )
        if best is None or best_score < self.threshold:
            return None
        return best

    def list_intents(self) -> list[dict]:
        return self.intents
