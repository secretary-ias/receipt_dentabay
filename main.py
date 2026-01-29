from pathlib import Path

from app.ui import run_app


def main() -> None:
    config_path = Path.cwd() / 'settings.json'
    run_app(config_path)


if __name__ == '__main__':
    main()
