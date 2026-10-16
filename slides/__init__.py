from collections.abc import Callable

from pptx import Presentation

from slides import binet, iterative, memoization, overview, person, problem, recurrence, recursion, title

SlideBuilder = Callable[[Presentation], None]

# Deck order: one entry per slide file, in presentation order.
SLIDES: list[SlideBuilder] = [
    title.build,
    person.build,
    problem.build,
    recurrence.build,
    recursion.build,
    memoization.build,
    iterative.build,
    binet.build,
    overview.build,
]
