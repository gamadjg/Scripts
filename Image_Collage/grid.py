"""
---functions---
get_grid_size():
 - get two inputs from user. How wide and tall you want the grid to be
 - if either of the two inputs are empty, set the empty value to 1

get_image():
 - take in input filepath for the image to be used.
 - if no file provided, use default file rocky.jpg

make_grid():
 - Create a simpleimge off of the file passed through the parameter
 - using the grid size x and y values (canvas_x, canvas_y) create a new canvas by using the size of the base image
    and multiplying copying it x and y times

layout when making the grid:
        00 -- 10 -- 20
        01 -- 11 -- 21
        02 -- 12 -- 22

quadrant_edit():
 - takes in the new canvas, original file, and which quadrant we are editing (x,y)
 - finds which quadrant within the new canvas, where will be editing
 - based on the quadrant x,y sum, determine what color scheme we will be offsetting the image by
 - for every column y and row x, get the pixel from the original image, offset the pixel color, then save at the new
    quadrant location.
 - quadrant_edit runs however however many times equal to the number of total quadrants, each time editing the
    next quadrant from left to right, top to bottom

"""
from simpleimage import SimpleImage


def get_grid_size():
    x = input('how wide do you want the grid to be:')
    y = input('how tall do you want the grid to be:')
    if x == '':
        x = 1
    if y == '':
        y = 1
    return int(x), int(y)


def get_image():
    image = input('enter image path or press enter to use default:')
    if image == '':
        image = 'rocky.jpg'
    return image


def make_grid(canvas_x, canvas_y, filename):
    base_image = SimpleImage(filename)

    # Get size of original image
    base_height = base_image.height
    base_width = base_image.width

    # make new 3x2 canvas with the base image size
    canvas = SimpleImage.blank(base_width*canvas_x, base_height*canvas_y)

    return canvas


def quadrant_edit(canvas, filename, x_quadrant, y_quadrant):
    base_image = SimpleImage(filename)

    # Get size of original image
    base_height = base_image.height
    base_width = base_image.width

    # Determine coordinates of where we will be editing on the new canvas, based on the original canvas size
    x_coordiante = base_width * int(x_quadrant)
    y_coordiante = base_height * int(y_quadrant)

    color_quadrant = x_quadrant + y_quadrant

    if color_quadrant == 0: # violet
        red_rgb = 1.5
        green_rgb = .7
        blue_rgb = 1.5
    elif color_quadrant == 1: # teal
        red_rgb = .7
        green_rgb = 1.5
        blue_rgb = 1.5
    elif color_quadrant == 2: # yellow
        red_rgb = 1.5
        green_rgb = 1.5
        blue_rgb = .7
    elif color_quadrant == 3: # teal
        red_rgb = .7
        green_rgb = 1.5
        blue_rgb = 1.5
    elif color_quadrant == 4: # violet
        red_rgb = 1.5
        green_rgb = .7
        blue_rgb = 1.5
    else:
        red_rgb = 1.5
        green_rgb = 1.5
        blue_rgb = 1.5

    for y in range(base_height):
        for x in range(base_width):
            opx = base_image.get_pixel(x, y)
            canvas_pixel = canvas.get_pixel(x_coordiante+x, y_coordiante+y)
            canvas_pixel.red = opx.red * red_rgb
            canvas_pixel.green = opx.green * green_rgb
            canvas_pixel.blue = opx.blue * blue_rgb

    return canvas

def main():
    x, y = get_grid_size()
    filename = get_image()
    grid = make_grid(x, y, filename)

    for width in range(x):
        for height in range(y):
            grid = quadrant_edit(grid, filename, width, height)
    grid.show()
    grid.save_file('rocky_6x6')

if __name__ == '__main__':
    main()
