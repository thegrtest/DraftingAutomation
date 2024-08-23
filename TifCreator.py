from PIL import Image, ImageDraw

# Create a blank white image
image = Image.new("RGB", (200, 200), "white")

# Initialize drawing context
draw = ImageDraw.Draw(image)

# Draw a simple shape (e.g., a rectangle and some text)
draw.rectangle([(50, 50), (150, 150)], outline="black", width=3)
draw.text((70, 90), "Test TIF", fill="black")

# Save the image as a .tif file
image.save("test_image.tif")

print("test_image.tif created successfully.")
