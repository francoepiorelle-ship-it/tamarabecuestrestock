from PIL import Image

# Abre tu imagen original
img = Image.open("logo.jpg")

# Guarda la imagen directamente en formato .ico con múltiples tamaños estándar
img.save("icono.ico", format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])

print("¡Ícono creado con éxito como 'icono.ico'!")