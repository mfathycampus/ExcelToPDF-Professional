"""Entry point of the packaged (PyInstaller) application.

    ExcelToPDF.exe                      -> desktop app
    ExcelToPDF.exe --cli <args>         -> command line (same as `x2p <args>`); use --report to get results
    ExcelToPDF.exe --x2p-worker excel <job.json> <result.json>
                                        -> internal: the Excel engine worker process
"""
import os
import sys


def _ensure_streams():
    # a windowed executable has no console: give print() somewhere harmless to go
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")


def main() -> int:
    _ensure_streams()
    args = sys.argv[1:]
    if args[:1] == ["--x2p-worker"]:
        if len(args) == 4 and args[1] == "excel":
            from x2p.engines import excel_worker
            excel_worker.main(args[2], args[3])
            return 0
        return 2
    if args[:1] == ["--cli"]:
        from x2p.cli import main as cli_main
        return cli_main(args[1:])
    from x2p.gui.app import main as gui_main
    return gui_main()


if __name__ == "__main__":
    raise SystemExit(main())
