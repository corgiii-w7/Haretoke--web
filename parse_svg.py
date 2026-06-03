import re
import json

def parse_svg(filename):
    with open(filename, 'r') as f:
        content = f.read()
    
    paths = re.findall(r'<path d="(.*?)" fill="(.*?)" transform="translate\((.*?)\)"', content)
    
    results = []
    for i, (d, fill, transform) in enumerate(paths):
        # find all numbers in d
        numbers = [float(x) for x in re.findall(r'[-+]?\d*\.\d+|\d+', d)]
        if not numbers:
            continue
        
        # d path might be relative or absolute, but we can roughly estimate bounds
        # actually there are C and M commands. Let's just grab all coordinates.
        xs = numbers[0::2]
        ys = numbers[1::2]
        
        # parse transform
        tx, ty = 0, 0
        if transform:
            t_parts = transform.split(',')
            if len(t_parts) == 2:
                tx, ty = float(t_parts[0]), float(t_parts[1])
                
        xs = [x + tx for x in xs]
        ys = [y + ty for y in ys]
        
        xmin, xmax = min(xs), max(xs)
        ymin, ymax = min(ys), max(ys)
        
        results.append({
            "index": i,
            "xmin": xmin, "xmax": xmax,
            "ymin": ymin, "ymax": ymax,
            "cx": (xmin + xmax)/2,
            "cy": (ymin + ymax)/2,
            "width": xmax - xmin,
            "height": ymax - ymin,
            "fill": fill
        })
        
    with open("svg_analysis.json", "w") as f:
        json.dump(results, f, indent=2)

parse_svg("haretoke_HP_image_home01.svg")
