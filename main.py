from pptx import Presentation
from pptx.util import Inches
import os
import json
import requests
import re
from pprint import pprint
from PIL import Image


text = {"outline": []}


def textReq(input, tokens):
    prompt = input
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer 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"
    }
    data = {
        "model": "text-davinci-003",
        "prompt": prompt,
        "temperature": 0,
        "max_tokens": tokens
    }
    response = requests.post("https://api.openai.com/v1/completions", headers=headers, data=json.dumps(data))
    return (response.json().get('choices')[0].get('text'))


def imageReq(input):
    # Set the headers for the request
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer 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"
    }

    # Set the data for the request
    data = {
        "prompt": input,
        "n": 1,
        "size": "512x512"
    }

    # Make the request to the OpenAI API
    response = requests.post("https://api.openai.com/v1/images/generations", headers=headers, data=json.dumps(data))

    # Check if the request was successful
    if response.status_code == 200:
        # Get the image data from the response
        image_data = response.json()["data"][0]["url"]

        # Download the image and save it to a file
        image = requests.get(image_data)
        open("image.jpg", "wb").write(image.content)
        print("Image saved to file.")
        return response.status_code
    else:
        print("Error: " + response.text)
        return response.status_code


# convert pixels to inches


print("What should the powerpoint be about?")
input = input("Enter your prompt: ")
prompt = "Background image for a presentation slide about "
imageReq(prompt + input)

prompt = "Write a title, subtitle, and an outline for an essay about "
res = textReq(prompt + input, 200)

print("\n Response text")
print(res)

title_match = re.search(r"Title: (.*?)\n", res)
subtitle_match = re.search(r"Subtitle: (.*?)\n", res)
outline_match = re.search(r"Outline:\n(.*)", res, re.DOTALL)

text["title"] = title_match.group(1) if title_match else None
text["subtitle"] = subtitle_match.group(1) if subtitle_match else None
outline = outline_match.group(1) if outline_match else None

print("\n Original outline")
print(outline)
# create presentation

pres = Presentation()

slide = pres.slides.add_slide(pres.slide_layouts[0])
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = text["title"]
subtitle.text = text["subtitle"]

left = top = Inches(0)
pic = slide.shapes.add_picture('image.jpg', left, top, width=pres.slide_width, height=pres.slide_height)

# This moves it to the background
slide.shapes._spTree.remove(pic._element)
slide.shapes._spTree.insert(2, pic._element)

# #Change the text formatting
# for shape in slide.shapes:
#     if not shape.has_text_frame:
#         continue
#         text_frame = shape.text_frame
#         # do things with the text frame


# pres.save('test.pptx')

# parse outline into parts

# get rid of the letters
# outline = re.sub("^[A-Z]{1,3}\.", "", outline, flags=re.MULTILINE)

# get rid of the trailing spaces
# outline = re.sub("^\s+", "", outline, flags=re.MULTILINE)

outlineParts = re.split(r"[IVX]+\.", outline.lstrip())

outlineParts[0] = outlineParts[0].lstrip()
print("Outline Parts: " + str(outlineParts))
i = -1
for part in outlineParts:
    if part.strip():
        text["outline"].append([])
        ++i
        lines = re.split(r"[A-Z]\.", part)
        print("Lines: " + str(lines))
        for line in lines:
            if line.strip():
                text["outline"][i].append(line.strip())




print("\nProcessed outline")
print(text["outline"])


# Iterate over the lists in the list
temp = "For a presentation about {}, write some bullet points for a slide titled \"{}\""


for i, list in enumerate(text["outline"]):
    for j, string in enumerate(list):
        string = string.lstrip()
        if j == 0 and len(list) > 1:
            print("Topic " + str(i) + ", Slide " + str(j) + ": " + string)
            slide = pres.slides.add_slide(pres.slide_layouts[0])
            title = slide.shapes.title
            title.text = string
        else:
            print("Topic " + str(i) + ", Slide " + str(j) + ": " + string)
            req = textReq(temp.format(input, string), 70).strip()
            print(req)

            bullets = req.replace("•", "")
            slide = pres.slides.add_slide(pres.slide_layouts[1])
            title_shape = slide.shapes.title
            body_shape = slide.shapes.placeholders[1]
            title_shape.text = string
            tf = body_shape.text_frame
            tf.text = bullets

# Save the presentation
pres.save('test.pptx')
