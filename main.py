from xml.dom import minidom
import PIL.Image

# This is a recursive method to go through an entire document to find the pic:pic elements.
def find_pic_elements(node):
    pic_elements = []
    for child_node in node.childNodes:
        if child_node.prefix == "pic" and child_node.localName == "pic":
            pic_elements.append(child_node)
        else:
            pic_elements.extend(find_pic_elements(child_node))
    return pic_elements

images = {}
# Open the file word/_rels/document.xml.rels, which defines an id for each image
image_id_document = minidom.parseString(open("word\_rels\document.xml.rels", "r", encoding="utf-8").read())
# Loop through the list of ids, only some of which are the images
for child_node in image_id_document.firstChild.childNodes:
    if child_node.localName == "Relationship" and child_node.getAttribute("Type").endswith("image"):
        id = child_node.getAttribute("Id")
        target = child_node.getAttribute("Target")
        # Map the image id to its file location
        images[id] = {"location": target}

# Open the main document, word/document.xml
document = minidom.parseString(open("word\document.xml", "r", encoding="utf-8").read())
# Find all pic:pic elements
pic_elements = find_pic_elements(document)
for element in pic_elements:
    id = ""
    src_rect = {}
    # Find the srcRect
    for child in element.childNodes:
        if child.localName == "blipFill":
            for child_node in child.childNodes:
                if child_node.localName == "blip":
                    id = child_node.getAttribute("r:embed")
                elif child_node.localName == "srcRect":
                    for side in "ltrb":
                        src_rect[side] = child_node.getAttribute(side)
    images[id]["src_rect"] = src_rect
    
# Find all images in images, crop them to the right size, and save them
# The crop values are in ST_Percentage, 1000ths of a percent
for id, info in images.items():
    image = PIL.Image.open("word/" + info["location"])
    width, height = image.size
    l = int(int(info["src_rect"]["l"]) / 100000 * width)
    t = int(int(info["src_rect"]["t"]) / 100000 * height)
    r = int(width - int(info["src_rect"]["r"]) / 100000 * width)
    b = int(height - int(info["src_rect"]["b"]) / 100000 * height)
    image = image.crop((l, t, r, b))
    image.save(info["location"])