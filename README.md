
# AI-Based PPT Generator

This project is an **AI-powered Streamlit web application** that automatically generates presentation slides from uploaded CSV files.  
It uses data analysis and a LLaMA 3.2 model to create meaningful titles, bullet points, charts, and detailed insights.

---

## 🚀 Features
- Upload a CSV file and select a target column.
- Choose visualization types: Scatter, Hexbin, Box, or Bar charts.
- AI-generated:
  - Slide titles and bullet points.
  - Introduction, comparison insights, detailed analysis, summaries, and conclusions.
- Minimum slide count customization.
- Generates a final downloadable `.odp` (OpenDocument Presentation) file.

---

## 🛠 Technologies Used
- **Python**
- **Streamlit** – for the web application.
- **pandas** – for data manipulation.
- **matplotlib** – for plotting charts.
- **python-pptx** – for creating slides.
- **Ollama (LLaMA 3.2 model)** – for AI-based text generation.
- **LibreOffice** – to convert `.pptx` to `.odp` format.

---

## How It Works
- Upload a CSV file.
- Select the target column and the desired chart type.
- Set minimum number of slides.
- Optionally, enter a custom prompt for LLaMA (like asking for a summary).
- Download the generated .odp report after processing.

---

## 📦 Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/your-repo-name.git
   cd your-repo-name
