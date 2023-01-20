from pptx import Presentation
from pptx.util import Inches
import os
import json
import requests
import re
from pprint import pprint
import logging

# Format the log message
logging.basicConfig(format='%(asctime)s %(levelname)s:%(name)s: %(message)s')
logger = logging.getLogger('dev')
logger.setLevel(logging.INFO)


text = {"outline": []}

def initTextReq(input):
    logger.info("Requesting title, subtitle, and outline...")
    prompt = "Write a title, subtitle, and an outline for an essay about {input}."
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer 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"
    }
    data = {
        "model": "text-davinci-003",
        "prompt": prompt.format(input=input),
        "temperature": 0,
        "max_tokens": 150
    }
    response = requests.post("https://api.openai.com/v1/completions", headers=headers, data=json.dumps(data))
    logger.debug("Response Text: " + response.json().get('choices')[0].get('text').strip())
    if response.status_code == 200:
        logger.info("Request successful.")
        return response.json().get('choices')[0].get('text')
    else:
        logger.error("Error: " + response.text)
        return response.json().get('choices')[0].get('text')

def quoteReq(input):
    logger.info("Requesting quote...")
    prompt = "Find a quote about {input}."
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer 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"
    }
    data = {
        "model": "text-davinci-003",
        "prompt": prompt.format(input=input),
        "temperature": 0,
        "max_tokens": 50
    }
    response = requests.post("https://api.openai.com/v1/completions", headers=headers, data=json.dumps(data))
    logger.debug("Response Text: " + response.json().get('choices')[0].get('text').strip())
    if response.status_code == 200:
        logger.info("Request successful.")
        return response.json().get('choices')[0].get('text').strip()
    else:
        logger.error("Error: " + response.text)
        return response.json().get('choices')[0].get('text').strip()


def analogyReq(input):
    logger.info("Requesting analogy...")
    prompt = "Write an analogy about {input}."
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer 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"
    }
    data = {
        "model": "text-davinci-003",
        "prompt": prompt.format(input=input),
        "temperature": 0,
        "max_tokens": 25
    }
    response = requests.post("https://api.openai.com/v1/completions", headers=headers, data=json.dumps(data))
    logger.debug("Response Text: " + response.json().get('choices')[0].get('text').strip())
    if response.status_code == 200:
        logger.info("Request successful.")
        return response.json().get('choices')[0].get('text').strip()
    else:
        logger.error("Error: " + response.text)
        return response.json().get('choices')[0].get('text').strip()

def bulletsReq(input,topic):
    logger.info("Requesting bullet points for " + topic + "...")
    prompt = "For a presentation about {input}, write some bullet points for a slide titled \"{topic}\""
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer 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"
    }
    data = {
        "model": "text-davinci-003",
        "prompt": prompt.format(input=input, topic=topic),
        "temperature": 0,
        "max_tokens": 70
    }
    response = requests.post("https://api.openai.com/v1/completions", headers=headers, data=json.dumps(data))
    logger.debug("Response Text: " + str(response.json().get('choices')[0].get('text').strip()))
    if response.status_code == 200:
        logger.info("Request successful.")
        return response.json().get('choices')[0].get('text').strip()
    else:
        logger.error("Error: " + response.text)
        return response.json().get('choices')[0].get('text').strip()


def backImageReq(input):
    logger.info("Requesting image...")
    prompt = "Background image for a presentation slide about {input}"
    # Set the headers for the request
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer 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"
    }

    # Set the data for the request
    data = {
        "prompt": prompt.format(input=input),
        "n": 1,
        "size": "512x512"
    }

    # Make the request to the OpenAI API
    response = requests.post("https://api.openai.com/v1/images/generations", headers=headers, data=json.dumps(data))
    logger.debug("Response: " + str(response))
    # Check if the request was successful
    if response.status_code == 200:
        # Get the image data from the response
        image_data = response.json()["data"][0]["url"]
        # Download the image and save it to a file
        image = requests.get(image_data)
        open("image.jpg", "wb").write(image.content)
        logger.info("Image saved to file.")
        return response.status_code
    else:
        logger.error("Error: " + response.text)
        return response.status_code

# convert pixels to inches

print("What should the powerpoint be about?")
input = input("Enter your prompt: ")
backImageReq(input)

initText = initTextReq(input)

title_match = re.search(r"Title: (.*?)\n", initText)
subtitle_match = re.search(r"Subtitle: (.*?)\n", initText)
outline_match = re.search(r"Outline:\n(.*)", initText, re.DOTALL)

text["title"] = title_match.group(1) if title_match else None
text["subtitle"] = subtitle_match.group(1) if subtitle_match else None
outline = outline_match.group(1) if outline_match else None

logger.debug("Title: " + text["title"])
logger.debug("Subtitle: " + text["subtitle"])
logger.debug("Outline: " + outline)

text["quote"] = quoteReq(input)
text["analogy"] = analogyReq(input)

logger.debug("Original Outline: " + outline)

# parse outline into parts
outlineParts = re.split(r"[IVX]+\.", outline.lstrip())

outlineParts[0] = outlineParts[0].lstrip()
logger.debug("Outline Parts: " + str(outlineParts))
i = -1
for part in outlineParts:
    if part.strip():
        text["outline"].append([])
        ++i
        lines = re.split(r"[A-Z]\.", part)
        logger.debug("Outline Lines: " + str(lines))
        for line in lines:
            if line.strip():
                text["outline"][i].append(line.strip())

logger.debug("Processed Outline: " + str(text["outline"]))


# create presentation

logger.info("Creating Presentation...")
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

logger.debug("Finished Title Slide")

#Analogy Slide
slide = pres.slides.add_slide(pres.slide_layouts[0])
title = slide.shapes.title
title.text = text["analogy"]

logger.debug("Finished Analogy Slide")

# Iterate over the lists in the list

for i, list in enumerate(text["outline"]):
    for j, string in enumerate(list):
        string = string.lstrip()
        if j == 0 and len(list) > 1:
            logger.info("Working on title slide " + str(j + 1) + "/" + str(len(list)) + " for topic " + str(i + 1) + "/" + str(len(text["outline"])))
            slide = pres.slides.add_slide(pres.slide_layouts[0])
            title = slide.shapes.title
            title.text = string
        else:
            logger.info("Working on content slide " + str(j + 1) + "/" + str(len(list)) + " for topic " + str(i + 1) + "/" + str(len(text["outline"])))
            req = bulletsReq(input, string).strip()

            bullets = req.replace("•", "")
            slide = pres.slides.add_slide(pres.slide_layouts[1])
            title_shape = slide.shapes.title
            body_shape = slide.shapes.placeholders[1]
            title_shape.text = string
            tf = body_shape.text_frame
            tf.text = bullets

logger.info("Finished Content Slides")

#Quote Slide
slide = pres.slides.add_slide(pres.slide_layouts[0])
title = slide.shapes.title
title.text = text["quote"]

logger.info("Finished Quote Slide")


# Save the presentation
pres.save('test.pptx')
logger.info("Finished Saving Presentation")

