# OpenAIPowerPointGenerator

A tool to generate PowerPoint presentations using OpenAI.

## Installation Instructions

Follow the steps below to set up and run the OpenAIPowerPointGenerator:

### 1. Clone the Repository

Begin by cloning the repository to your local machine:

```bash
git clone https://github.com/noahyoungs/OpenAIPowerPointGenerator.git
cd OpenAIPowerPointGenerator
```

### 2. Create a Python 3.8 Environment

Ensure you have Python 3.8 installed. Create a virtual environment and activate it:

**For macOS and Linux:**

```bash
python3 -m venv venv
source venv/bin/activate
```

**For Windows:**

```bash
python -m venv venv
.\venv\Scripts\activate
```

Once activated, install the required packages using `requirements.txt`:

```bash
pip install -r requirements.txt
```

### 3. Set up the .env File

The repository contains a `.env.template` file. You need to save this as `.env` and then update it with your OpenAI API key:

First, make a copy and rename:

```bash
cp .env.template .env
```

Then, open the `.env` file in a text editor of your choice and replace `YOUR_API_KEY` with your actual OpenAI API key.

### 4. Run the Application

Once everything is set up, run the application using:

```bash
python main.py
```

---

