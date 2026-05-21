"""
core/fuzzy_matcher.py
─────────────────────
Utility for approximate string matching used when aligning invoice item names
to master-file records.  Uses rapidfuzz (C-extension; very fast).
"""

from __future__ import annotations
from typing import Optional
from rapidfuzz import fuzz, process


# ─── Default similarity thresholds ───────────────────────────────────────────
THRESHOLD_HIGH   = 90   # Near-certain match (SKU typo, extra space)
THRESHOLD_MEDIUM = 78   # Probable match (word order, abbreviation)
THRESHOLD_LOW    = 65   # Possible match – flagged as "needs review"


class FuzzyMatcher:
    """
    Wraps rapidfuzz utilities for finding the closest matching string from
    a candidate list.  Prefers token_sort_ratio (insensitive to word order)
    and falls back to partial_ratio for substring matches.
    """

    def __init__(self, threshold: int = THRESHOLD_MEDIUM):
        self.threshold = threshold

    # ── Public API ─────────────────────────────────────────────────────────

    def best_match(
        self,
        query: str,
        candidates: list[str],
        threshold: Optional[int] = None,
    ) -> tuple[Optional[str], int]:
        """
        Return (best_candidate, score) where score is 0-100.
        Returns (None, 0) when no candidate meets the threshold.

        Strategy:
          1. Exact match (case-insensitive)  → score 100
          2. token_sort_ratio                → handles word-order differences
          3. partial_ratio                   → handles substring cases
        """
        if not candidates:
            return None, 0

        limit = threshold if threshold is not None else self.threshold
        q_norm = query.strip().lower()

        # 1. Exact match
        for c in candidates:
            if c.strip().lower() == q_norm:
                return c, 100

        # 2. token_sort_ratio  (best for multi-word names)
        result_sort = process.extractOne(
            query,
            candidates,
            scorer=fuzz.token_sort_ratio,
        )

        # 3. partial_ratio  (best for truncated names)
        result_partial = process.extractOne(
            query,
            candidates,
            scorer=fuzz.partial_ratio,
        )

        # Pick the higher-scoring result from the two strategies
        candidates_scored = [r for r in [result_sort, result_partial] if r is not None]
        if not candidates_scored:
            return None, 0

        best = max(candidates_scored, key=lambda r: r[1])
        best_name, best_score, *_ = best

        if best_score >= limit:
            return best_name, best_score
        return None, best_score

    def rank_matches(
        self,
        query: str,
        candidates: list[str],
        n: int = 5,
    ) -> list[tuple[str, int]]:
        """Return the top-n (name, score) tuples, best first."""
        results = process.extract(
            query, candidates,
            scorer=fuzz.token_sort_ratio, limit=n)
        return [(r[0], r[1]) for r in results]

    @staticmethod
    def normalise(text: str) -> str:
        """Lowercase, strip, collapse internal whitespace."""
        import re
        return re.sub(r"\s+", " ", text.strip().lower())


# ─── Module-level convenience singleton ───────────────────────────────────────
_default_matcher = FuzzyMatcher()

def best_match(query: str, candidates: list[str],
               threshold: int = THRESHOLD_MEDIUM) -> tuple[Optional[str], int]:
    return _default_matcher.best_match(query, candidates, threshold)
