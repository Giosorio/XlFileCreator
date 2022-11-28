


def set_color_code(color_code):
    def text_color(text):
        return f"\033[{color_code}{text}\033["

    return text_color



black = set_color_code('90m')
red = set_color_code('91m')
green = set_color_code('92m')
yellow = set_color_code('93m')
blue = set_color_code('94m')
magenta = set_color_code('95m')
cyan = set_color_code('96m')
white = set_color_code('97m')