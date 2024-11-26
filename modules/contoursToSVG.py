import numpy as np
import matplotlib.pyplot as plt
from skimage.draw import bezier_curve
from scipy.interpolate import splprep, splev

# Load the image
image = plt.imread("test2.png")

# Convert the image to grayscale
gray = np.dot(image[..., :3], [0.2989, 0.5870, 0.1140])

# Threshold the image to obtain the contours
threshold = gray > 0.5

# Find the contours in the thresholded image
rows, cols = np.where(threshold)
points = np.column_stack((cols, rows))

# Use scipy's spline functions to fit a Bezier curve to the contours
tck, u = splprep(points.T, u=None, s=0.0, per=1)
curve = splev(np.linspace(0, 1, 25), tck)

# Use skimage's bezier curve function to generate the curve
bezier = np.array(bezier_curve(*curve, number_of_curves=1))

# Plot the curve
fig, ax = plt.subplots()
ax.imshow(image)
ax.plot(bezier[0], bezier[1], "r-", lw=5)
plt.show()

# Save the curve as an SVG file
with open("bezier.svg", "w") as f:
    f.write("<svg xmlns='http://www.w3.org/2000/svg' width='1200' height='800'>\n")
    f.write("  <path d='M ")
    for i in range(bezier.shape[1]):
        f.write(f"{bezier[0, i]} {bezier[1, i]} ")
    f.write("' stroke='red' stroke-width='5' fill='none' />\n")
    f.write("</svg>")
