#! python3
# A multi-clipboard program.

TEXT = {
    "invoice": """\nI've processed the invoice.\n\nCheers,\n\nMatt\n\n""",
    "expense": """\nI've processed your expense claim.\n\nCheers,\n\nMatt\n\n""",
    "upsell": """Would you consider making this a monthly donation?""",
}

import sys, pyperclip, pyautogui, time

if len(sys.argv) < 2:
    print("Usage: py mclip.py [keyphrase] - copy phrase text")
    sys.exit()

keyphrase = sys.argv[1]  # first command line arg is the keyphrase

if keyphrase in TEXT:
    pyperclip.copy(TEXT[keyphrase])
    print("Text for " + keyphrase + " copied to clipboard.")
    fw = pyautogui.getActiveWindow()
    if (
        len(
            pyautogui.getWindowsWithTitle(
                "MERIEUX NUTRISCIENCES CORPORATION Mail - Google Chrome"
            )
        )
        >= 1
    ):
        pyautogui.getWindowsWithTitle(
            "MERIEUX NUTRISCIENCES CORPORATION Mail - Google Chrome"
        )[0].activate()
        fw.close()
else:
    print("There is no text for " + keyphrase)
