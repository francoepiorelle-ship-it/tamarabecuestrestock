import webview


def iniciar_programa():
    # URL de tu aplicación en Streamlit (forzando tema claro si deseas)
    url = "https://tamarabecuestrestock.streamlit.app/?auth=ok&embed_options=light_theme"

    webview.create_window(
        title="Gestión Tamara B - Control de Stock",
        url=url,
        width=1280,
        height=850,
        resizable=True,
        min_size=(800, 600),
    )
    webview.start()


if __name__ == "__main__":
    iniciar_programa()