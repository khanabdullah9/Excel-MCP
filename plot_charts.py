import matplotlib.pyplot as plt
from typing import Sequence, Optional

def plot_line_chart(y_axis: Sequence[float], x_axis: Optional[Sequence[float]] = None) -> plt.Figure:
    """Generates a line chart figure."""
    fig, ax = plt.subplots()
    
    if x_axis is not None and len(x_axis) > 0:
        ax.plot(x_axis, y_axis)
    else:
        ax.plot(y_axis)

    return fig

def plot_pie_chart(labels: Sequence[str], sizes: Sequence[float]) -> plt.Figure:
    """Generates a pie chart figure."""
    fig, ax = plt.subplots()
    
    ax.pie(sizes, labels = labels, autopct='%1.1f%%')

    return fig