import json
from pptx.util import Inches


def create_json_dicts():
    try:
        mapping = {"pic_mapping": {"Red": [[Inches(1.85), Inches(0.75), Inches(0.5), Inches(0.5), False],
                                           [Inches(3.85), Inches(0.75), Inches(0.5), Inches(0.5), False],
                                           [Inches(5.85), Inches(0.75), Inches(0.5), Inches(0.5), False],
                                           [Inches(7.85), Inches(0.75), Inches(0.5), Inches(0.5), False]],
                                   "Green": [[Inches(1.85), Inches(2.75), Inches(0.5), Inches(0.5), False],
                                             [Inches(3.85), Inches(2.75), Inches(0.5), Inches(0.5), False],
                                             [Inches(5.85), Inches(2.75), Inches(0.5), Inches(0.5), False],
                                             [Inches(7.85), Inches(2.75), Inches(0.5), Inches(0.5), False]],
                                   "Blue": [[Inches(1.85), Inches(4.75), Inches(0.5), Inches(0.5), False],
                                            [Inches(3.85), Inches(4.75), Inches(0.5), Inches(0.5), False],
                                            [Inches(5.85), Inches(4.75), Inches(0.5), Inches(0.5), False],
                                            [Inches(7.85), Inches(4.75), Inches(0.5), Inches(0.5), False]]},
                   "pic_list": {"Red": [None] * 4,
                                "Green": [None] * 4,
                                "Blue": [None] * 4},
                   "line_mapping": {"Red": [[Inches(2.1), Inches(1.5)], [Inches(4.1), Inches(1.5)],
                                            [Inches(6.1), Inches(1.5)], [Inches(8.1), Inches(1.5)]],
                                    "Green": [[Inches(2.1), Inches(3.5)], [Inches(4.1), Inches(3.5)],
                                              [Inches(6.1), Inches(3.5)], [Inches(8.1), Inches(3.5)]],
                                    "Blue": [[Inches(2.1), Inches(5.5)], [Inches(4.1), Inches(5.5)],
                                             [Inches(6.1), Inches(5.5)], [Inches(8.1), Inches(5.5)]]},
                   "lines": []}
        with open('data.json', 'w') as outfile:
            json.dump(mapping, outfile)
        return
    except Exception as e:
        print(e)
