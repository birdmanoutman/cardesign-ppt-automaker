import math
import numpy as np
import matplotlib.pyplot as plt

from skimage.draw import (line, polygon, disk,
                          circle_perimeter,
                          ellipse, ellipse_perimeter,
                          bezier_curve)


fig, (ax1, ax2) = plt.subplots(ncols=2, nrows=1, figsize=(10, 6))


img = np.zeros((500, 500, 3), dtype=np.double)

# Bezier curve
rr, cc = bezier_curve(70, 100, 10, 10, 150, 100, 1)
img[rr, cc, :] = (1, 0, 0)

ax1.imshow(img)
ax1.set_title('No anti-aliasing')
ax1.axis('off')
plt.show()