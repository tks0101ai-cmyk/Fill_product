from contract_filler.gui import App


def test_app_has_two_tabs_with_expected_titles():
    app = App()
    try:
        assert app.notebook.index("end") == 2
        assert app.notebook.tab(0, "text") == "Tạo hợp đồng mua bán"
        assert app.notebook.tab(1, "text") == "Tạo hợp đồng nguyên tắc"
    finally:
        app.destroy()
