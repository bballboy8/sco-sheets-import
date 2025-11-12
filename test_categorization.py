"""
Tests for the business categorization logic in sco_import.py

This test suite verifies that keywords are matched as standalone words
rather than substrings within other words.
"""

import os
import pytest

# Mock environment variables before importing sco_import to avoid gspread dependency
os.environ.setdefault("SHEET_ID", "test_sheet_id")
os.environ.setdefault("SEEN_IDS_SHEET_ID", "test_seen_ids_sheet_id")
os.environ.setdefault("GOOGLE_CREDENTIALS", '{"type": "service_account"}')

# Now import the functions we need to test
from sco_import import is_business, _matches_with_boundaries


class TestMatchesWithBoundaries:
    """Test the _matches_with_boundaries helper function."""
    
    def test_matches_standalone_word_at_start(self):
        """Keyword at the start of the string should match."""
        assert _matches_with_boundaries("CAR REPAIR", "CAR") is True
        assert _matches_with_boundaries("CAR ", "CAR") is True
    
    def test_matches_standalone_word_in_middle(self):
        """Keyword in the middle with spaces should match."""
        assert _matches_with_boundaries("ABC CAR REPAIR", "CAR") is True
        assert _matches_with_boundaries("MY CAR SHOP", "CAR") is True
    
    def test_matches_standalone_word_at_end(self):
        """Keyword at the end should match."""
        assert _matches_with_boundaries("REPAIR CAR", "CAR") is True
        assert _matches_with_boundaries("REPAIR CAR ", "CAR") is True
    
    def test_does_not_match_substring_in_word(self):
        """Keyword as substring within another word should NOT match."""
        assert _matches_with_boundaries("CARLSON", "CAR") is False
        assert _matches_with_boundaries("SCAR", "CAR") is False
        assert _matches_with_boundaries("CARSEN", "CAR") is False
        assert _matches_with_boundaries("CARPET", "CAR") is False
    
    def test_does_not_match_partial_word_boundary(self):
        """Should not match if only one boundary is correct."""
        assert _matches_with_boundaries("CARL", "CAR") is False  # 'L' is alphanumeric
        assert _matches_with_boundaries("SCAR ", "CAR") is False  # 'S' is alphanumeric
    
    def test_matches_with_punctuation(self):
        """Should match with punctuation boundaries."""
        assert _matches_with_boundaries("CAR, INC", "CAR") is True
        assert _matches_with_boundaries("CAR-REPAIR", "CAR") is True
        assert _matches_with_boundaries("(CAR)", "CAR") is True
    
    def test_case_insensitive_handled_by_caller(self):
        """Note: This function receives uppercase strings from is_business."""
        assert _matches_with_boundaries("CAR REPAIR", "CAR") is True
        assert _matches_with_boundaries("car repair", "car") is True


class TestIsBusiness:
    """Test the is_business function with various owner names."""
    
    # Test cases for "CAR" keyword (the example from the user)
    def test_car_matches_standalone(self):
        """CAR should match when it's a standalone word."""
        assert is_business("CAR REPAIR LLC") is True
        assert is_business("ABC CAR SHOP") is True
        assert is_business("CAR SALES INC") is True
        assert is_business("MY CAR DEALERSHIP") is True
    
    def test_car_does_not_match_substring(self):
        """CAR should NOT match when it's part of another word."""
        # Use names without other business terms to test CAR specifically
        assert is_business("CARLSON") is False
        assert is_business("SCAR") is False
        assert is_business("CARSEN") is False
        assert is_business("CARLSON FAMILY") is False
        # Note: "CARPET WORLD" would return True because "CARPET" is itself a business term,
        # but "CAR" keyword doesn't match "CARPET" - verified by the other test cases above
    
    # Test cases for "CARS" keyword
    def test_cars_matches_standalone(self):
        """CARS should match when it's a standalone word."""
        assert is_business("CARS R US") is True
        assert is_business("ABC CARS INC") is True
    
    def test_cars_does_not_match_substring(self):
        """CARS should NOT match when it's part of another word."""
        # Use name without other business terms
        assert is_business("CARSEN") is False
    
    # Test cases for "TRUCK" keyword
    def test_truck_matches_standalone(self):
        """TRUCK should match when it's a standalone word."""
        assert is_business("TRUCK REPAIR") is True
        assert is_business("ABC TRUCK SHOP") is True
    
    def test_truck_does_not_match_substring(self):
        """TRUCK should NOT match when it's part of another word."""
        # Use name without other business terms
        assert is_business("STRUCK") is False
    
    # Test cases for "BOAT" keyword
    def test_boat_matches_standalone(self):
        """BOAT should match when it's a standalone word."""
        assert is_business("BOAT SALES") is True
        assert is_business("ABC BOAT SHOP") is True
    
    def test_boat_does_not_match_substring(self):
        """BOAT should NOT match when it's part of another word."""
        # Use name without other business terms
        assert is_business("STEAMBOAT") is False
    
    # Test cases for terms that can appear anywhere (A, THE, OF, AND)
    def test_common_words_as_standalone(self):
        """Common words should match as standalone words anywhere in the name."""
        assert is_business("A COMPANY") is True
        assert is_business("THE CORPORATION") is True
        assert is_business("A SMITH") is True  # "A" as standalone word
        assert is_business("THE JONES") is True  # "THE" as standalone word
        assert is_business("ABC A SMITH") is True  # "A" as standalone word in middle
        assert is_business("ABC THE JONES") is True  # "THE" as standalone word in middle
        # But should not match as substrings
        assert is_business("ABC") is False  # "A" is part of "ABC"
        assert is_business("THEATER") is False  # "THE" is part of "THEATER"
    
    # Test cases for space-required terms (OF, AND)
    def test_space_required_terms(self):
        """Terms that require spaces on both sides."""
        assert is_business("COMPANY OF AMERICA") is True
        assert is_business("SMITH AND JONES") is True
        assert is_business("OF AMERICA") is True
        # Note: "AND " requires a space after, so "SMITH AND JONES" should match
        # because "AND " is followed by "JONES" (word boundary)
        assert is_business("SMITH AND JONES") is True
    
    # Test cases for partial match terms (should still work as substring)
    def test_partial_match_terms(self):
        """Partial match terms should still match as substrings."""
        assert is_business("COMMUNICATIONS INC") is True  # Contains "COMMUNICAT"
        assert is_business("CONSTRUCTION LLC") is True  # Contains "CONSTRUCTIO"
        assert is_business("ENTERTAINMENT CORP") is True  # Contains "ENTERTAINME"
        assert is_business("MANUFACTURING CO") is True  # Contains "MANUFACTUR"
        assert is_business("PERFORMANCE AUTO") is True  # Contains "PERFORMANC"
        assert is_business("SMITH & JONES") is True  # Contains "&"
    
    # Test cases for common business terms
    def test_common_business_terms(self):
        """Common business terms should match correctly."""
        assert is_business("ABC LLC") is True
        assert is_business("XYZ INC") is True
        assert is_business("COMPANY CORP") is True
        assert is_business("HOSPITAL ASSOCIATION") is True
        assert is_business("UNIVERSITY FOUNDATION") is True
    
    # Edge cases
    def test_empty_string(self):
        """Empty strings should return False."""
        assert is_business("") is False
        assert is_business("   ") is False
    
    def test_none_value(self):
        """None should return False."""
        assert is_business(None) is False
    
    def test_individual_names(self):
        """Individual names should return False."""
        assert is_business("JOHN SMITH") is False
        assert is_business("MARY JONES") is False
        assert is_business("ROBERT CARLSON") is False  # Contains "CAR" but not as standalone
        assert is_business("JANE DOE") is False
    
    def test_case_insensitivity(self):
        """Function should be case-insensitive."""
        assert is_business("car repair") is True
        assert is_business("Car Repair") is True
        assert is_business("CAR REPAIR") is True
        # Use names without other business terms
        assert is_business("carlson") is False
        assert is_business("Carlson") is False
        assert is_business("CARLSON") is False
    
    def test_multiple_keywords(self):
        """Should match if any keyword matches."""
        assert is_business("CAR REPAIR LLC") is True  # Matches both CAR and LLC
        assert is_business("TRUCK SALES INC") is True  # Matches both TRUCK and INC
    
    def test_keyword_with_punctuation(self):
        """Keywords should match even with punctuation boundaries."""
        assert is_business("CAR, INC") is True
        assert is_business("CAR-REPAIR") is True
        assert is_business("(CAR) SALES") is True


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

