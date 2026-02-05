"""
Unit tests for the Deterministic Formula Engine.
Tests formula generation for various statistical tasks.
"""

import pytest
from tools.formula_engine import (
    FormulaEngine,
    FormulaType,
    ColumnMapping,
    FormulaResult,
    create_formula_engine
)


class TestColumnMapping:
    """Tests for ColumnMapping dataclass."""
    
    def test_column_mapping_data_range(self):
        """Test data_range property generates correct Excel reference."""
        mapping = ColumnMapping(
            name="Age",
            letter="B",
            index=2,
            data_start_row=2,
            data_end_row=101
        )
        
        assert mapping.data_range == "'00_RAW_DATA_LOCKED'!B2:B101"
    
    def test_column_mapping_absolute_range(self):
        """Test absolute_range property generates correct Excel reference."""
        mapping = ColumnMapping(
            name="Score",
            letter="C",
            index=3,
            data_start_row=2,
            data_end_row=50
        )
        
        assert mapping.absolute_range == "'00_RAW_DATA_LOCKED'!$C$2:$C$50"


class TestFormulaEngine:
    """Tests for FormulaEngine class."""
    
    @pytest.fixture
    def sample_columns(self):
        """Sample column names for testing."""
        return ["ID", "Age", "Score", "Group", "Rating"]
    
    @pytest.fixture
    def engine(self, sample_columns):
        """Create a FormulaEngine instance for testing."""
        return FormulaEngine(columns=sample_columns, n_rows=100)
    
    def test_engine_initialization(self, engine, sample_columns):
        """Test engine initializes with correct column mappings."""
        assert len(engine.column_map) == len(sample_columns)
        assert engine.n_rows == 100
        assert engine.data_start_row == 2
        assert engine.data_end_row == 101
    
    def test_get_column(self, engine):
        """Test get_column returns correct ColumnMapping."""
        col = engine.get_column("Age")
        
        assert col is not None
        assert col.name == "Age"
        assert col.letter == "B"
        assert col.index == 2
    
    def test_get_column_not_found(self, engine):
        """Test get_column returns None for unknown column."""
        col = engine.get_column("NonExistent")
        assert col is None
    
    def test_get_column_range(self, engine):
        """Test get_column_range returns correct range string."""
        range_str = engine.get_column_range("Score")
        assert range_str == "'00_RAW_DATA_LOCKED'!C2:C101"
    
    def test_get_column_range_raises_for_unknown(self, engine):
        """Test get_column_range raises ValueError for unknown column."""
        with pytest.raises(ValueError, match="Column 'Unknown' not found"):
            engine.get_column_range("Unknown")
    
    def test_generate_descriptive_formulas(self, engine):
        """Test descriptive statistics formula generation."""
        results = engine.generate_descriptive_formulas(
            col_name="Age",
            output_col=2,
            output_start_row=5
        )
        
        assert len(results) == 11
        
        formula_types = [r.formula_type for r in results]
        assert FormulaType.COUNT in formula_types
        assert FormulaType.MEAN in formula_types
        assert FormulaType.STDEV in formula_types
        assert FormulaType.SKEWNESS in formula_types
        
        count_formula = next(r for r in results if r.formula_type == FormulaType.COUNT)
        assert "COUNT(" in count_formula.formula
        assert "'00_RAW_DATA_LOCKED'!B2:B101" in count_formula.formula
    
    def test_generate_correlation_formula(self, engine):
        """Test correlation formula generation."""
        formula = engine.generate_correlation_formula("Age", "Score")
        
        assert formula.startswith("=CORREL(")
        assert "'00_RAW_DATA_LOCKED'!B2:B101" in formula
        assert "'00_RAW_DATA_LOCKED'!C2:C101" in formula
    
    def test_generate_ttest_formula(self, engine):
        """Test t-test formula generation."""
        formula = engine.generate_ttest_formula("Age", "Score", tails=2, test_type=2)
        
        assert formula.startswith("=T.TEST(")
        assert ",2,2)" in formula
    
    def test_generate_grouped_mean_formula(self, engine):
        """Test grouped mean (AVERAGEIF) formula generation."""
        formula = engine.generate_grouped_mean_formula(
            value_col="Score",
            group_col="Group",
            group_value="A"
        )
        
        assert formula.startswith("=AVERAGEIF(")
        assert '"A"' in formula
    
    def test_generate_grouped_count_formula(self, engine):
        """Test grouped count (COUNTIF) formula generation."""
        formula = engine.generate_grouped_count_formula(
            group_col="Group",
            group_value=1
        )
        
        assert formula.startswith("=COUNTIF(")
        assert "1" in formula
    
    def test_generate_missing_analysis_formulas(self, engine):
        """Test missing data analysis formula generation."""
        results = engine.generate_missing_analysis_formulas(
            col_names=["Age", "Score"],
            output_start_row=5
        )
        
        assert len(results) == 6
        
        valid_n = [r for r in results if r.formula_type == FormulaType.VALID_N]
        missing = [r for r in results if r.formula_type == FormulaType.MISSING]
        
        assert len(valid_n) == 2
        assert len(missing) == 2


class TestCreateFormulaEngine:
    """Tests for the factory function."""
    
    def test_create_formula_engine(self):
        """Test factory function creates engine correctly."""
        columns = ["A", "B", "C"]
        engine = create_formula_engine(columns=columns, n_rows=50)
        
        assert isinstance(engine, FormulaEngine)
        assert engine.n_rows == 50
        assert len(engine.column_map) == 3


class TestFormulaResult:
    """Tests for FormulaResult dataclass."""
    
    def test_formula_result_creation(self):
        """Test FormulaResult can be created with all fields."""
        result = FormulaResult(
            cell="B5",
            formula="=AVERAGE(A2:A100)",
            label="Mean",
            formula_type=FormulaType.MEAN
        )
        
        assert result.cell == "B5"
        assert result.formula == "=AVERAGE(A2:A100)"
        assert result.label == "Mean"
        assert result.formula_type == FormulaType.MEAN
