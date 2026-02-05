"""
Unit tests for the Deterministic Verification Module.
Tests statistical verification against Python ground truth.
"""

import pytest
import numpy as np
import pandas as pd
from pathlib import Path
import tempfile

from tools.verification import (
    StatisticalVerifier,
    VerificationStatus,
    VerificationCheck,
    VerificationResult,
    DEFAULT_TOLERANCE,
    STATISTICAL_TOLERANCE,
    generate_verification_report
)


class TestStatisticalVerifier:
    """Tests for StatisticalVerifier class."""
    
    @pytest.fixture
    def sample_data(self):
        """Create sample DataFrame for testing."""
        np.random.seed(42)
        return pd.DataFrame({
            'age': np.random.normal(35, 10, 100),
            'score': np.random.normal(75, 15, 100),
            'rating': np.random.randint(1, 6, 100),
            'group': np.random.choice(['A', 'B'], 100)
        })
    
    @pytest.fixture
    def verifier(self, sample_data):
        """Create StatisticalVerifier instance."""
        return StatisticalVerifier(sample_data)
    
    def test_compute_descriptives(self, verifier, sample_data):
        """Test descriptive statistics computation."""
        stats = verifier.compute_descriptives('age')
        
        assert 'count' in stats
        assert 'mean' in stats
        assert 'std' in stats
        assert 'min' in stats
        assert 'max' in stats
        assert 'skewness' in stats
        assert 'kurtosis' in stats
        
        assert stats['count'] == 100
        assert abs(stats['mean'] - sample_data['age'].mean()) < 1e-10
        assert abs(stats['std'] - sample_data['age'].std(ddof=1)) < 1e-10
    
    def test_compute_descriptives_with_missing(self, sample_data):
        """Test descriptives with missing values."""
        sample_data.loc[0:9, 'age'] = np.nan
        verifier = StatisticalVerifier(sample_data)
        
        stats = verifier.compute_descriptives('age')
        
        assert stats['count'] == 90
        assert stats['missing'] == 10
    
    def test_compute_correlation(self, verifier, sample_data):
        """Test correlation computation."""
        r = verifier.compute_correlation('age', 'score')
        
        expected_r = sample_data['age'].corr(sample_data['score'])
        assert abs(r - expected_r) < 1e-10
    
    def test_compute_correlation_insufficient_data(self):
        """Test correlation with insufficient data."""
        df = pd.DataFrame({'a': [1, 2], 'b': [np.nan, np.nan]})
        verifier = StatisticalVerifier(df)
        
        r = verifier.compute_correlation('a', 'b')
        assert np.isnan(r)
    
    def test_compute_ttest(self, verifier):
        """Test t-test computation."""
        t, p = verifier.compute_ttest('age', 'score', paired=False)
        
        assert isinstance(t, float)
        assert isinstance(p, float)
        assert 0 <= p <= 1
    
    def test_compute_frequency(self, verifier, sample_data):
        """Test frequency computation."""
        freq = verifier.compute_frequency('group')
        
        assert 'A' in freq
        assert 'B' in freq
        assert freq['A'] + freq['B'] == 100
    
    def test_compute_cronbach_alpha(self):
        """Test Cronbach's alpha computation."""
        np.random.seed(42)
        items = pd.DataFrame({
            'item1': np.random.randint(1, 6, 50),
            'item2': np.random.randint(1, 6, 50),
            'item3': np.random.randint(1, 6, 50)
        })
        items['item2'] = items['item1'] + np.random.normal(0, 0.5, 50)
        items['item3'] = items['item1'] + np.random.normal(0, 0.5, 50)
        
        verifier = StatisticalVerifier(items)
        alpha = verifier.compute_cronbach_alpha(['item1', 'item2', 'item3'])
        
        assert isinstance(alpha, float)
        assert 0 <= alpha <= 1
    
    def test_compute_shapiro_wilk(self, verifier):
        """Test Shapiro-Wilk test computation."""
        w, p = verifier.compute_shapiro_wilk('age')
        
        assert isinstance(w, float)
        assert isinstance(p, float)
        assert 0 <= w <= 1
        assert 0 <= p <= 1
    
    def test_compute_cohens_d(self):
        """Test Cohen's d computation."""
        np.random.seed(42)
        group1 = pd.Series(np.random.normal(100, 15, 50))
        group2 = pd.Series(np.random.normal(110, 15, 50))
        
        df = pd.DataFrame({'g1': group1, 'g2': group2})
        verifier = StatisticalVerifier(df)
        d = verifier.compute_cohens_d(group1, group2)
        
        assert isinstance(d, float)
        assert abs(d) < 5


class TestVerificationCheck:
    """Tests for VerificationCheck dataclass."""
    
    def test_within_tolerance_pass(self):
        """Test within_tolerance returns True when difference is small."""
        check = VerificationCheck(
            check_name="mean",
            expected_value=10.0,
            actual_value=10.0001,
            tolerance=0.001,
            status=VerificationStatus.PASS,
            cell_reference="B5"
        )
        
        assert check.within_tolerance is True
        assert check.difference < 0.001
    
    def test_within_tolerance_fail(self):
        """Test within_tolerance returns False when difference is large."""
        check = VerificationCheck(
            check_name="mean",
            expected_value=10.0,
            actual_value=11.0,
            tolerance=0.001,
            status=VerificationStatus.FAIL,
            cell_reference="B5"
        )
        
        assert check.within_tolerance is False
        assert check.difference == 1.0
    
    def test_difference_none_when_actual_none(self):
        """Test difference returns None when actual_value is None."""
        check = VerificationCheck(
            check_name="mean",
            expected_value=10.0,
            actual_value=None,
            tolerance=0.001,
            status=VerificationStatus.FAIL,
            cell_reference="B5"
        )
        
        assert check.difference is None
        assert check.within_tolerance is False


class TestVerificationResult:
    """Tests for VerificationResult dataclass."""
    
    def test_passed_checks_count(self):
        """Test passed_checks property counts correctly."""
        result = VerificationResult(
            task_id="1.1",
            sheet_name="TEST",
            status=VerificationStatus.PASS,
            checks=[
                VerificationCheck("a", 1.0, 1.0, 0.01, VerificationStatus.PASS, "A1"),
                VerificationCheck("b", 2.0, 2.0, 0.01, VerificationStatus.PASS, "A2"),
                VerificationCheck("c", 3.0, 4.0, 0.01, VerificationStatus.FAIL, "A3"),
            ]
        )
        
        assert result.passed_checks == 2
        assert result.failed_checks == 1
        assert result.total_checks == 3
    
    def test_pass_rate_calculation(self):
        """Test pass_rate calculation."""
        result = VerificationResult(
            task_id="1.1",
            sheet_name="TEST",
            status=VerificationStatus.PASS,
            checks=[
                VerificationCheck("a", 1.0, 1.0, 0.01, VerificationStatus.PASS, "A1"),
                VerificationCheck("b", 2.0, 2.0, 0.01, VerificationStatus.PASS, "A2"),
                VerificationCheck("c", 3.0, 4.0, 0.01, VerificationStatus.FAIL, "A3"),
                VerificationCheck("d", 4.0, 4.0, 0.01, VerificationStatus.PASS, "A4"),
            ]
        )
        
        assert result.pass_rate == 75.0
    
    def test_pass_rate_empty_checks(self):
        """Test pass_rate returns 0 for empty checks."""
        result = VerificationResult(
            task_id="1.1",
            sheet_name="TEST",
            status=VerificationStatus.SKIP,
            checks=[]
        )
        
        assert result.pass_rate == 0.0


class TestGenerateVerificationReport:
    """Tests for report generation."""
    
    def test_report_generation(self):
        """Test verification report is generated correctly."""
        results = [
            VerificationResult(
                task_id="1.1",
                sheet_name="TEST_SHEET",
                status=VerificationStatus.PASS,
                checks=[
                    VerificationCheck("mean", 10.0, 10.0, 0.01, VerificationStatus.PASS, "B5"),
                ],
                formula_coverage=85.0
            )
        ]
        
        report = generate_verification_report(results)
        
        assert "VERIFICATION REPORT" in report
        assert "Task 1.1" in report
        assert "TEST_SHEET" in report
        assert "Formula Coverage: 85.0%" in report
        assert "PASS" in report or "✓" in report


class TestTolerances:
    """Tests for tolerance constants."""
    
    def test_default_tolerance(self):
        """Test default tolerance is reasonable."""
        assert DEFAULT_TOLERANCE == 1e-6
    
    def test_statistical_tolerance(self):
        """Test statistical tolerance is reasonable."""
        assert STATISTICAL_TOLERANCE == 1e-4
