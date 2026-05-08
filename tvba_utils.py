"""Utility functions for unit conversions and text processing.

Corresponds to VBA FormatModule.bas:
  - ConvertSizeToPoints
  - CentimetersToPoints
  - CleanParaText
"""

_SIZE_LABEL_TO_POINTS = {
    "初号": 42.0,
    "小初": 36.0,
    "一号": 26.0,
    "小一": 24.0,
    "二号": 22.0,
    "小二": 18.0,
    "三号": 16.0,
    "小三": 15.0,
    "四号": 14.0,
    "小四": 12.0,
    "五号": 10.5,
    "小五": 9.0,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5.0,
}

_POINTS_TO_SIZE_LABEL = {v: k for k, v in _SIZE_LABEL_TO_POINTS.items()}


def size_label_to_points(label: str) -> float:
    """Convert Chinese size label or numeric string to points."""
    label = label.strip()
    if label in _SIZE_LABEL_TO_POINTS:
        return _SIZE_LABEL_TO_POINTS[label]
    # Try parsing as numeric, e.g. "14" or "14pt"
    num = label.replace("pt", "").strip()
    try:
        return float(num)
    except ValueError:
        raise ValueError(f"Unknown size label: {label!r}")


def points_to_size_label(points: float, tolerance: float = 0.5) -> str:
    """Convert points back to Chinese size label, or return 'Xpt' string."""
    for pt, label in _POINTS_TO_SIZE_LABEL.items():
        if abs(points - pt) <= tolerance:
            return label
    return f"{points:g}pt"


def cm_to_points(cm: float) -> float:
    """Convert centimeters to points. 1 cm = 28.3465 points."""
    return cm * 28.3465


def points_to_cm(points: float) -> float:
    """Convert points to centimeters."""
    return points / 28.3465


def clean_para_text(text: str) -> str:
    """Strip whitespace and normalize paragraph text for matching.

    Corresponds to VBA CleanParaText.
    """
    return text.strip().replace("\r", "").replace("\n", "")
