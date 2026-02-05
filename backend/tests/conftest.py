"""
Pytest configuration and fixtures for PhD Survey Analyzer tests.
"""

import pytest
import numpy as np
import pandas as pd
from pathlib import Path
import tempfile
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))


@pytest.fixture
def sample_survey_data():
    """
    Create sample survey data for testing.
    Mimics a typical Likert-scale survey with demographics.
    """
    np.random.seed(42)
    n = 100
    
    data = {
        'respondent_id': range(1, n + 1),
        'age': np.random.randint(18, 65, n),
        'gender': np.random.choice(['Male', 'Female', 'Other'], n, p=[0.48, 0.48, 0.04]),
        'education': np.random.choice(['High School', 'Bachelor', 'Master', 'PhD'], n),
        'q1_satisfaction': np.random.randint(1, 6, n),
        'q2_satisfaction': np.random.randint(1, 6, n),
        'q3_satisfaction': np.random.randint(1, 6, n),
        'q4_engagement': np.random.randint(1, 6, n),
        'q5_engagement': np.random.randint(1, 6, n),
        'q6_engagement': np.random.randint(1, 6, n),
        'q7_loyalty': np.random.randint(1, 6, n),
        'q8_loyalty': np.random.randint(1, 6, n),
        'income': np.random.choice([30000, 50000, 75000, 100000, 150000], n),
        'open_feedback': [f"Feedback {i}" if np.random.random() > 0.3 else np.nan for i in range(n)]
    }
    
    return pd.DataFrame(data)


@pytest.fixture
def sample_numeric_data():
    """Create sample numeric-only data for statistical tests."""
    np.random.seed(42)
    n = 50
    
    return pd.DataFrame({
        'var1': np.random.normal(100, 15, n),
        'var2': np.random.normal(50, 10, n),
        'var3': np.random.normal(75, 20, n),
        'group': np.random.choice(['Control', 'Treatment'], n)
    })


@pytest.fixture
def temp_excel_file(sample_survey_data):
    """Create a temporary Excel file with sample data."""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        sample_survey_data.to_excel(f.name, index=False)
        yield Path(f.name)


@pytest.fixture
def column_list():
    """Standard column list for formula engine tests."""
    return ['ID', 'Age', 'Score', 'Group', 'Rating', 'Q1', 'Q2', 'Q3']


@pytest.fixture
def scale_items():
    """Sample scale items for reliability tests."""
    return ['q1_satisfaction', 'q2_satisfaction', 'q3_satisfaction']
