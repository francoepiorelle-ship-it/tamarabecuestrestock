import webview

def iniciar_programa():
    # Añadimos el parámetro para forzar el tema claro de Streamlit si la aplicación lo soporta
    url = "https://tamarabecuestrestock.streamlit.app/?auth=ok&embed_options=light_theme"
    
    window = webview.create_window(
        title="Gestión Tamara B - Control de Stock",
        url=url,
        width=1280,
        height=850,
        resizable=True,
        min_size=(800, 600)
    )
    
    webview.start(private_mode=False)

if __name__ == "__main__":
    iniciar_programa()