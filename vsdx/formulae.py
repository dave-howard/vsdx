import math
from vsdx import Shape


def width_x_1(shape: Shape):
    return shape.width


def width_x_0(shape: Shape):
    return 0


def middle_x(shape: Shape):
    # (BeginX+EndX)/2
    return (shape.begin_x + shape.end_x) / 2


def middle_y(shape: Shape):
    # (BeginY+EndY)/2
    return (shape.begin_y + shape.end_y) / 2


def center_x(shape: Shape):
    # Width*0.5
    return shape.width * 0.5


def center_y(shape: Shape):
    # Height*0.5
    return shape.height * 0.5


def diag_width(shape: Shape):
    # SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)
    width = shape.end_x - shape.begin_x
    height = shape.end_y - shape.begin_y
    return math.sqrt(width ** 2 + height ** 2)


def angle(shape: Shape):
    # ATAN2(EndY-BeginY,EndX-BeginX)
    w = shape.end_x-shape.begin_x
    h = shape.end_y-shape.begin_y
    return math.atan2(w, h)


def width(shape: Shape):
    return shape.end_x-shape.begin_x


def height(shape: Shape):
    return shape.end_y-shape.begin_y


# map func text to functions
func_map = {
    'Width*1': width_x_1,
    'Width*0': width_x_0,
    '(BeginX+EndX)/2': middle_x,
    '(BeginY+EndY)/2': middle_y,
    'Width*0.5': center_x,
    'Height*0.5': center_y,
    'SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)': diag_width,
    'ATAN2(EndY-BeginY,EndX-BeginX)': angle,
    'GUARD((BeginX+EndX)/2)': middle_x,
    'GUARD((BeginY+EndY)/2)': middle_y,
    'GUARD(Width*0.5)': center_x,
    'GUARD(Height*0.5)': center_y,
    'GUARD(EndX-BeginX)': width,
    'GUARD(EndY-BeginY)': height,
}


def calc_value(shape: Shape, func_text: str):
    f = func_map.get(func_text)
    if f:
        return f(shape=shape)
    elif shape.page.vis.debug:
        # show any non-matching formulae
        print(f"calc_value(func_text='{func_text}') no method found")
