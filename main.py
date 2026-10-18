import platform
import subprocess

from fastmcp import FastMCP

mcp = FastMCP("LiveWordEditor")

OS = platform.system()


def _run_applescript(script: str, *args: str) -> str:
    """Runs AppleScript, passing values as arguments (no string escaping needed)."""
    result = subprocess.run(
        ["osascript", "-e", script, *args],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript Error: {result.stderr.strip()}")
    return result.stdout.strip()


_MAC_TEXT_SCRIPT = """
on run argv
    set theText to item 1 of argv
    tell application "Microsoft Word"
        activate
        type paragraph selection
        type text selection text theText
    end tell
end run
"""

_MAC_EQUATION_SCRIPT = """
on run argv
    set theLatex to item 1 of argv
    tell application "Microsoft Word"
        activate
        type paragraph selection
        set startPos to selection start of selection
        type text selection text theLatex
        set selection start of selection to startPos
    end tell
    delay 0.2
    tell application "System Events"
        keystroke "=" using {control down} -- Word: insert/convert to equation
        delay 0.2
        key code 124 -- Right arrow: leave the equation box
    end tell
end run
"""


def _mac_insert_text(text: str) -> None:
    _run_applescript(_MAC_TEXT_SCRIPT, text)


def _mac_insert_equation(latex: str) -> None:
    _run_applescript(_MAC_EQUATION_SCRIPT, latex)


def _word_app():
    import win32com.client as win32

    try:
        return win32.GetActiveObject("Word.Application")
    except Exception as exc:
        raise RuntimeError(
            "Microsoft Word is not running. Open Word and a document first."
        ) from exc


def _win_insert_text(text: str) -> None:
    sel = _word_app().Selection
    sel.TypeParagraph()
    sel.TypeText(text)


def _win_insert_equation(latex: str) -> None:
    app = _word_app()
    doc = app.ActiveDocument
    sel = app.Selection

    sel.TypeParagraph()
    start = sel.Start
    sel.TypeText(latex)
    rng = doc.Range(start, sel.End)

    eq_range = doc.OMaths.Add(rng)
    omath = eq_range.OMaths(1)
    omath.BuildUp()

    end = omath.Range.End
    sel.SetRange(end, end)


def _unsupported() -> RuntimeError:
    return RuntimeError(
        f"Unsupported OS '{OS}'. This server supports macOS and Windows."
    )


@mcp.tool()
def get_platform() -> str:
    """Returns the operating system this Word server is running on."""
    return {"Darwin": "macOS", "Windows": "Windows"}.get(OS, OS)


@mcp.tool()
def insert_text(text: str) -> str:
    """Inserts plain text on a new line at the cursor in Microsoft Word (macOS or Windows)."""
    if OS == "Darwin":
        _mac_insert_text(text)
    elif OS == "Windows":
        _win_insert_text(text)
    else:
        raise _unsupported()
    return "Text successfully inserted."


@mcp.tool()
def insert_equation(latex_string: str) -> str:
    """
    Inserts a LaTeX string on a new line as a native Word equation
    (macOS or Windows). Word's equation input should be set to LaTeX.
    """
    latex = latex_string.strip()
    if OS == "Darwin":
        _mac_insert_equation(latex)
    elif OS == "Windows":
        _win_insert_equation(latex)
    else:
        raise _unsupported()
    return f"Equation '{latex}' successfully built in Word."


if __name__ == "__main__":
    mcp.run()
