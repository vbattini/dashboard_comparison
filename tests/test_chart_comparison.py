import tempfile
import unittest
from pathlib import Path

from Text_Extraction___ import (
    category_match_score,
    compare_extractions,
    sort_and_format_chart_items,
    split_chart_values,
    synchronize_identical_pair_extractions,
)


class ChartComparisonTests(unittest.TestCase):
    def test_series_values_split_without_breaking_thousands(self):
        self.assertEqual(split_chart_values("2,208", 2), ["2,208"])
        self.assertEqual(split_chart_values("96, 61", 2), ["96", "61"])
        self.assertEqual(
            split_chart_values("1,200, 2,400", 2),
            ["1,200", "2,400"],
        )

    def test_chart_title_provides_series_bindings_when_missing(self):
        items = sort_and_format_chart_items(
            [
                {
                    "section": "Chart",
                    "item_name": "Cases Registered vs Cases Resolved - Sunday",
                    "value": "96, 61",
                    "series_name": None,
                }
            ]
        )

        self.assertEqual(
            [(item["series_name"], item["value"]) for item in items],
            [("Cases Registered", "96"), ("Cases Resolved", "61")],
        )

    def test_comparison_matches_values_across_extraction_shapes(self):
        trendence_item = {
            "section": "Chart",
            "item_name": "Cases Registered vs Cases Resolved - Sunday",
            "value": "96,61",
            "series_name": "Cases Registered, Cases Resolved",
            "visual_type": "clustered_bar_chart",
        }
        spartnash_item = {
            "section": "Chart",
            "item_name": "Cases Registered vs Cases Resolved - Sunday",
            "value": "96, 61",
            "series_name": None,
            "visual_type": "clustered_bar_chart",
        }

        result = compare_extractions(
            {"items": [trendence_item]},
            {"items": [spartnash_item]},
        )

        self.assertEqual(result["comparison_items"][0]["status"], "Match")
        self.assertEqual(
            result["comparison_items"][0]["item_name"],
            "Cases Registered vs Cases Resolved - Sunday",
        )
        self.assertEqual(
            result["comparison_items"][0]["trendence_series"],
            {"cases registered": "96", "cases resolved": "61"},
        )

    def test_comparison_detects_swapped_series_values(self):
        baseline = {
            "section": "Chart",
            "item_name": "Cases Registered vs Cases Resolved - Sunday",
            "value": "96, 61",
            "series_name": None,
            "visual_type": "clustered_bar_chart",
        }
        swapped = {
            **baseline,
            "value": "61, 96",
        }

        result = compare_extractions(
            {"items": [baseline]},
            {"items": [swapped]},
        )

        self.assertEqual(result["comparison_items"][0]["status"], "Different")

    def test_unmatched_item_in_shared_visual_is_uncertain(self):
        trendence = {
            "section": "Chart",
            "item_name": "Cases by Segment - Distribution",
            "value": "10",
            "visual_type": "bar_chart",
        }
        spartnash = {
            "section": "Chart",
            "item_name": "Cases by Segment - Retail",
            "value": "20",
            "visual_type": "bar_chart",
        }

        result = compare_extractions(
            {"items": [trendence]},
            {"items": [spartnash]},
        )

        self.assertEqual(
            {row["status"] for row in result["comparison_items"]},
            {"Uncertain"},
        )

    def test_item_in_missing_visual_remains_one_sided(self):
        result = compare_extractions(
            {
                "items": [
                    {
                        "section": "Chart",
                        "item_name": "Trendence-only chart",
                        "value": "10",
                        "visual_type": "bar_chart",
                    }
                ]
            },
            {"items": []},
        )

        self.assertEqual(
            result["comparison_items"][0]["status"],
            "Trendence Only",
        )

    def test_truncated_stacked_bar_category_matches_and_uses_full_label(self):
        series = "Time to Offer Acceptance, Offer Acceptance to Joining"
        trendence = {
            "section": "Chart",
            "item_name": "Time to Offer Acceptance vs Offer Acceptance to Joining - Transp...",
            "value": "27.6, 8.1",
            "series_name": series,
            "visual_type": "stacked_bar_chart",
        }
        spartnash = {
            **trendence,
            "item_name": "Time to Offer Acceptance vs Offer Acceptance to Joining - Transportation",
        }

        result = compare_extractions(
            {"items": [trendence]},
            {"items": [spartnash]},
        )
        row = result["comparison_items"][0]

        self.assertEqual(row["status"], "Match")
        self.assertEqual(
            row["item_name"],
            "Time to Offer Acceptance vs Offer Acceptance to Joining - Transportation",
        )
        self.assertEqual(
            row["trendence_series"],
            {
                "time to offer acceptance": "27.6",
                "offer acceptance to joining": "8.1",
            },
        )

    def test_category_match_does_not_drop_qualifiers(self):
        self.assertLess(
            category_match_score("External Caller", "External Caller (VOE Only)"),
            0.78,
        )

    def test_identical_images_share_the_more_complete_extraction(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            trendence_image = root / "trendence.png"
            spartnash_image = root / "spartnash.png"
            image_bytes = b"identical dashboard image"
            trendence_image.write_bytes(image_bytes)
            spartnash_image.write_bytes(image_bytes)

            trendence = {
                "items": [
                    {
                        "section": "series_value",
                        "item_name": "A-Category",
                        "category": "Category",
                        "series_name": "A",
                        "value": "12",
                        "_visual_title": "Chart",
                    },
                    {
                        "section": "series_value",
                        "item_name": "B-Category",
                        "category": "Category",
                        "series_name": "B",
                        "value": "8",
                        "_visual_title": "Chart",
                    },
                ]
            }
            spartnash = {
                "items": [
                    {
                        "section": "series_value",
                        "item_name": "A-Category",
                        "category": "Category",
                        "series_name": "A",
                        "value": "12",
                        "_visual_title": "Chart",
                    }
                ]
            }
            trendence_extractions = {14: trendence}
            spartnash_extractions = {14: spartnash}

            synchronized = synchronize_identical_pair_extractions(
                {14: trendence_image},
                {14: spartnash_image},
                trendence_extractions,
                spartnash_extractions,
            )

        self.assertEqual(synchronized, [14])
        self.assertEqual(trendence_extractions, spartnash_extractions)
        result = compare_extractions(
            trendence_extractions[14],
            spartnash_extractions[14],
        )
        self.assertTrue(
            all(row["status"] == "Match" for row in result["comparison_items"])
        )

    def test_duplicate_extraction_notes_are_written_once(self):
        result = compare_extractions(
            {"items": [], "notes": ["Same image was extracted"]},
            {"items": [], "notes": ["Same image was extracted"]},
        )

        self.assertEqual(result["comparison_notes"], ["Same image was extracted"])


if __name__ == "__main__":
    unittest.main()
