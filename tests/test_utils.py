import pytest
from tvba_utils import size_label_to_points, points_to_size_label, cm_to_points, points_to_cm


class TestSizeLabelToPoints:
    def test_init_size(self):
        assert size_label_to_points("初号") == 42.0

    def test_xiaosi_size(self):
        assert size_label_to_points("小四") == 12.0

    def test_wuhao_size(self):
        assert size_label_to_points("五号") == 10.5

    def test_numeric_string(self):
        assert size_label_to_points("14") == 14.0

    def test_numeric_with_pt(self):
        assert size_label_to_points("14pt") == 14.0

    def test_unknown_label_raises(self):
        with pytest.raises(ValueError, match="Unknown size label"):
            size_label_to_points("不存在")


class TestPointsToSizeLabel:
    def test_exact_match(self):
        assert points_to_size_label(12.0) == "小四"

    def test_close_match(self):
        assert points_to_size_label(12.1) == "小四"

    def test_no_close_match_returns_string(self):
        assert points_to_size_label(13.0) == "13pt"


class TestCmToPoints:
    def test_one_cm(self):
        assert cm_to_points(1.0) == pytest.approx(28.3465, rel=1e-4)

    def test_zero_cm(self):
        assert cm_to_points(0.0) == 0.0


class TestPointsToCm:
    def test_one_cm_in_points(self):
        assert points_to_cm(28.3465) == pytest.approx(1.0, rel=1e-4)

    def test_zero_points(self):
        assert points_to_cm(0.0) == 0.0
