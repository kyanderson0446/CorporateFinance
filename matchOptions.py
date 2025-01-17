import re
from typing import List, Tuple, Optional
from fuzzywuzzy import fuzz
import jellyfish
from difflib import SequenceMatcher

class FacilityNameMatcher:
    @staticmethod
    def compare_matchers(name1: str, name2: str, threshold: int = 85) -> dict:
        """
        Compare different text matching algorithms for facility name matching.

        Args:
            name1 (str): First facility name
            name2 (str): Second facility name
            threshold (int): Score threshold for considering a match

        Returns:
            dict: Scores from different matching algorithms
        """

        def normalize_name(name):
            # Convert to lowercase and remove common words/punctuation
            name = name.lower()
            name = re.sub(r'\b(center|centre|rehabilitation|rehab|nursing|care|skilled|health|healthcare|post|acute|post-acute|memory|community|facility)\b', '', name)
            name = re.sub(r'[^\w\s]', '', name)
            return ' '.join(name.split())

        # Normalize names
        norm1 = normalize_name(name1)
        norm2 = normalize_name(name2)

        return {
            # Token-based comparisons (better for word rearrangement)
            "token_sort_ratio": fuzz.token_sort_ratio(norm1, norm2),
            "token_set_ratio": fuzz.token_set_ratio(norm1, norm2),

            # Sequence-based comparisons (better for typos/spelling variations)
            "levenshtein_distance": jellyfish.levenshtein_distance(norm1, norm2),
            "jaro_winkler": int(jellyfish.jaro_winkler_similarity(norm1, norm2) * 100),

            # Hybrid approaches
            "sequence_matcher": int(SequenceMatcher(None, norm1, norm2).ratio() * 100),
            "metaphone": jellyfish.metaphone(norm1) == jellyfish.metaphone(norm2),
        }

    @staticmethod
    def match_facility_name(
        row: str,
        choices: List[str],
        threshold: int = 85,
        debug: bool = False
    ) -> Tuple[Optional[str], float]:
        """
        Enhanced facility name matcher using multiple algorithms.

        Args:
            row: The facility name to match
            choices: List of possible facility names to match against
            threshold: Minimum score to consider a match
            debug: If True, print detailed matching information

        Returns:
            tuple: (best_match, confidence_score) or (None, 0) if no match
        """
        # Input validation
        if not isinstance(choices, (list, tuple)):
            raise ValueError(f"choices must be a list or tuple, got {type(choices)}")

        if not choices:
            return None, 0

        if not isinstance(row, str) or len(row.strip()) == 0:
            return None, 0

        # Filter out any non-string or empty choices
        valid_choices = [c for c in choices if isinstance(c, str) and len(c.strip()) > 0]

        if debug:
            print(f"\nMatching facility: {row}")
            print(f"Number of valid choices: {len(valid_choices)}")

        best_match = None
        best_score = 0

        for choice in valid_choices:
            if len(choice) <= 1:  # Skip single characters
                continue

            scores = FacilityNameMatcher.compare_matchers(row, choice, threshold)

            # Calculate composite score
            token_score = max(scores['token_sort_ratio'], scores['token_set_ratio'])
            sequence_score = scores['jaro_winkler']

            # Weighted average (emphasize token-based matching)
            composite_score = (token_score * 0.7 + sequence_score * 0.3)

            if scores['metaphone']:
                composite_score = min(100, composite_score + (10 if scores['metaphone'] else 0))

            if debug:
                print(f"\nComparing with: {choice}")
                print(f"Token score: {token_score}")
                print(f"Sequence score: {sequence_score}")
                print(f"Composite score: {composite_score}")
                print(f"All scores: {scores}")

            # Early rejection for clearly different names
            if scores['levenshtein_distance'] > min(len(row), len(choice)) / 2:
                if debug:
                    print(f"Rejected due to high Levenshtein distance: {scores['levenshtein_distance']}")
                continue

            if composite_score > best_score and composite_score >= threshold:
                best_score = composite_score
                best_match = choice

        if debug:
            print(f"\nBest match: {best_match}")
            print(f"Best score: {best_score}")

        return best_match, best_score
