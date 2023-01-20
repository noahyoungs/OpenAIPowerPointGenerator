from pptx import Presentation
from pptx.util import Inches
import os
import json
import requests
import re
from pprint import pprint

print("What should the powerpoint be about?")
input = input("Enter your prompt: ")

prompt = "Write a title, subtitle, and an outline for an essay about " + input

headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer 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"
}

data = {
    "model": "text-davinci-003",
    "prompt": prompt,
    "temperature": 0,
    "max_tokens": 100
}

response = requests.post("https://api.openai.com/v1/completions", headers=headers, data=json.dumps(data))
text = response.json().get('choices')[0].get('text')
print(text)

title_match = re.search(r"Title: (.*?)\n", text)
subtitle_match = re.search(r"Subtitle: (.*?)\n", text)
outline_match = re.search(r"Outline:\n(.*)", text, re.DOTALL)

title = title_match.group(1) if title_match else None
subtitle = subtitle_match.group(1) if subtitle_match else None
outline = outline_match.group(1) if outline_match else None

print("Title: ", title)
print("Subtitle: ", subtitle)
print("Outline: ", outline)

# parts = text.split(":")
#
# title = parts[0]
# subtitle = parts[1]
#
# pres = Presentation()
#
# Layout = pres.slide_layouts[0]
# first_slide = pres.slides.add_slide(Layout)
#
# first_slide.shapes.title.text = title
#
# first_slide.placeholders[1].text = subtitle
#
# pres.save("First_presentation.pptx")
#

