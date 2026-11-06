from research_mcp.windows import WindowInfo, is_suspicious_zotero_window


def test_zotero_window_classifier_ignores_main_and_launcher_windows():
    assert not is_suspicious_zotero_window(
        WindowInfo(1, 1, "paper - Zotero", "MozillaWindowClass", True)
    )


def test_zotero_window_classifier_flags_dialogs():
    assert is_suspicious_zotero_window(WindowInfo(1, 1, "Warning", "#32770", True))
    assert is_suspicious_zotero_window(
        WindowInfo(2, 1, "Zotero Integration Error", "MozillaDialogClass", True)
    )
    # Word integration field prompts can use the executable path as their title.
    assert is_suspicious_zotero_window(
        WindowInfo(3, 1, r"C:\Program Files\Zotero\zotero.exe", "#32770", True)
    )
