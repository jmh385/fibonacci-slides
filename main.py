from pathlib import Path

from pptx import Presentation

from slides import SLIDES

OUTPUT_PATH = Path("target/fibonacci-slides.pptx")


def main() -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    for build in SLIDES:
        build(prs)
    prs.save(str(OUTPUT_PATH))
    print(f"Saved {OUTPUT_PATH} ({len(SLIDES)} slides)")


if __name__ == "__main__":
    main()
