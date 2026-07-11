from io import BytesIO

from matplotlib.figure import Figure

_CHART_KINDS = ("bar", "line", "pie")

_FIGURE_SIZE = (8.0, 4.5)
_FIGURE_DPI = 150


def render_chart(
    kind: str,
    categories: list[str],
    series: dict[str, list[float]],
    title: str | None,
) -> bytes:

    if kind not in _CHART_KINDS:
        raise ValueError(
            f"Chart kind '{kind}' not supported. "
            f"Use one of: {', '.join(_CHART_KINDS)}."
        )

    if not categories or not series:
        raise ValueError("categories and series must not be empty.")

    for name, values in series.items():
        if len(values) != len(categories):
            raise ValueError(
                f"Series '{name}' has {len(values)} values, "
                f"but there are {len(categories)} categories."
            )

    if kind == "pie" and len(series) > 1:
        raise ValueError("A pie chart takes exactly one series.")

    figure = Figure(figsize=_FIGURE_SIZE, dpi=_FIGURE_DPI, layout="constrained")
    axes = figure.add_subplot()

    if kind == "bar":
        bar_width = 0.8 / len(series)

        for offset, (name, values) in enumerate(series.items()):
            positions = [
                i - 0.4 + bar_width * (offset + 0.5)
                for i in range(len(categories))
            ]
            axes.bar(positions, values, width=bar_width, label=name)

        axes.set_xticks(range(len(categories)), categories)

    elif kind == "line":
        for name, values in series.items():
            axes.plot(categories, values, marker="o", label=name)

    else:
        (values,) = series.values()
        axes.pie(values, labels=categories, autopct="%1.1f%%")

    if kind != "pie":
        axes.grid(axis="y", alpha=0.3)
        axes.legend()

    if title is not None:
        axes.set_title(title)

    buffer = BytesIO()
    figure.savefig(buffer, format="png")

    return buffer.getvalue()
