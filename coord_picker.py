import sys
from typing import List, Tuple

from PIL import Image, ImageTk
import tkinter as tk


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: python coord_picker.py <image_path>")
        sys.exit(1)

    path = sys.argv[1]
    img = Image.open(path)

    root = tk.Tk()
    root.title("Click top-left and bottom-right of chat area")

    tk_img = ImageTk.PhotoImage(img)
    canvas = tk.Canvas(root, width=tk_img.width(), height=tk_img.height())
    canvas.pack()
    canvas.create_image(0, 0, anchor=tk.NW, image=tk_img)

    points: List[Tuple[int, int]] = []

    def on_click(event):
        points.append((event.x, event.y))
        print(f"Clicked: {event.x},{event.y}")
        if len(points) == 2:
            (x1, y1), (x2, y2) = points
            left, top = min(x1, x2), min(y1, y2)
            right, bottom = max(x1, x2), max(y1, y2)
            print(f"CROP_RECT={left},{top},{right},{bottom}")
            root.after(300, root.destroy)

    canvas.bind("<Button-1>", on_click)
    root.mainloop()


if __name__ == "__main__":
    main()
