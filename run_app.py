from pathlib import Path
import sys

# Keep the source layout runnable without requiring a package-install step.
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from printcase.app import main

if __name__ == "__main__":
    main()
