
# ... début du script (imports, fonctions, etc.)

# Fonction convert_ecmg_to_inches supposée existante dans le script précédent :
def convert_ecmg_to_inches(value, axis="x"):
    value = float(value)
    if axis == "x":
        return (value + 1.299) * (30.48 / 149.351)
    else:
        return (value + 10.917) * (18.543 / 152.838)

# → Calcul précis des positions PowerPoint (en inches)
if design_el is not None and design_el.attrib.get("left") not in [None, ""]:
    left = Inches(convert_ecmg_to_inches(design_el.attrib["left"], "x"))
else:
    left = Inches(convert_ecmg_to_inches(style.get("left", "0"), "x"))

if design_el is not None and design_el.attrib.get("top") not in [None, ""]:
    top = Inches(convert_ecmg_to_inches(design_el.attrib["top"], "y"))
else:
    top = Inches(convert_ecmg_to_inches(style.get("top", "0"), "y"))

if design_el is not None and design_el.attrib.get("width") not in [None, ""]:
    width = Inches(convert_ecmg_to_inches(design_el.attrib["width"], "x"))
else:
    width = Inches(convert_ecmg_to_inches(style.get("width", "5"), "x"))

if design_el is not None and design_el.attrib.get("height") not in [None, ""]:
    height = Inches(convert_ecmg_to_inches(design_el.attrib["height"], "y"))
else:
    height = Inches(convert_ecmg_to_inches(style.get("height", "1"), "y"))

print(f"Ajout box at → top={round(top.inches*2.54, 2)} cm, left={round(left.inches*2.54, 2)} cm, width={round(width.inches*2.54, 2)} cm, height={round(height.inches*2.54, 2)} cm")
