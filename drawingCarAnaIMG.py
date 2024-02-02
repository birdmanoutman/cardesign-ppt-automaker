import matplotlib.pyplot as plt
import numpy as np
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from PIL import Image
import pandas as pd
import os
from matplotlib.font_manager import FontProperties
# new test


# Load the data from the Excel file
df = pd.read_excel('carPriceAndSeg/mpv_data_with_images.xlsx')

# Update the 'Image_Path' column in the DataFrame with the correct paths
df['Image_Path'] = df['Image_Path'].apply(lambda x: os.path.join('carPriceAndSeg/car-pic', os.path.basename(x)))


# Function to resize an image
def resize_image(image_path, target_height):
    image = Image.open(image_path)
    original_width, original_height = image.size
    aspect_ratio = original_width / original_height
    target_width = int(target_height * aspect_ratio)
    resized_image = image.resize((target_width, target_height))
    return np.array(resized_image)


# Function to calculate text size
def get_text_size(text, font_properties, fig, ax):
    font = FontProperties(family=font_properties)  # Use FontProperties to set the font family
    text_artist = ax.text(0, 0, text, fontproperties=font)
    fig.canvas.draw()
    bbox = text_artist.get_window_extent(renderer=fig.canvas.get_renderer())
    text_artist.remove()
    return bbox.width, bbox.height


# Create the figure and axes
fig, ax = plt.subplots(figsize=(16, 8))

# Plot the bars and calculate the center positions for the images
bar_centers = {}
for index, row in df.iterrows():
    bar = ax.bar(row['Wheelbase'], row['Max Price'] - row['Min Price'], bottom=row['Min Price'], color='grey', edgecolor='black', alpha=0.5)
    bar_centers[row['Car_name']] = bar[0].get_x() + bar[0].get_width() / 2

# Plot the images at the center of the bars
for index, row in df.iterrows():
    image = resize_image(row['Image_Path'], target_height=50)
    imagebox = OffsetImage(image, zoom=0.5)
    ab = AnnotationBbox(imagebox, (bar_centers[row['Car_name']], row['Min Price'] + (row['Max Price'] - row['Min Price']) / 2), frameon=False)
    ax.add_artist(ab)

# Initialize the list to store the positions of the text
texts = []
text_positions = []

# Set a minimum distance that we want to keep between the image and the text
min_distance = 10  # This is an arbitrary number, you may need to adjust it

# Add text labels for each car
for index, row in df.iterrows():
    # Estimate text size
    text = f"{row['Car_name']}\n{row['Dimension']}\n{row['Positioning']}"
    text_width, text_height = get_text_size(text, 'Arial Unicode MS', fig, ax)  # Update font name as needed

    # Position the text above the image with a minimum distance
    text_x = bar_centers[row['Car_name']]
    text_y = row['Max Price'] + text_height / fig.dpi + min_distance / fig.dpi  # Convert pixels to points
    text_artist = ax.text(text_x, text_y, text, ha='center', va='bottom', fontproperties=FontProperties(family='Arial Unicode MS'))


    texts.append(text_artist)
    text_positions.append((text_x, text_y))

# Setting the labels and title
ax.set_xlabel('轴距 (mm)')
ax.set_ylabel('价格 (万元)')
ax.set_title('车辆轴距与价格区间关系')

# Show the plot
plt.tight_layout()
plt.show()
