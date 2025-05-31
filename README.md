# Document-Processing-RAG-Pipeline

![PDF Extraction Pipeline](https://github.com/user-attachments/assets/fa1ae41f-f0b3-445c-913d-1b03e69a8172)

**Before you begin**, please ensure you've **configured the following requirements within the `main()` function** (typically in `main.py` or a dedicated configuration section) in STEP 1:
* **LLMWhispererV2 API Key:** PDF data extraction tool.
* **Google GenAI API Key:** Retrieval Augmented Generation (RAG) engine and Embeddings that generates answers.
* **`Dataformodel.txt`:** Knowledge base or contextual data for your RAG engine.
* **`INPUTPDF`:** PDF document from which questions will be extracted.
* **Adobe PDF Services Client ID:** For conversion of the generated `DOCX file` into a `final PDF document`.
* **Adobe PDF Services Client Secret:** The corresponding secret key for authentication with Adobe PDF Services.

#### How to Use this Tool?

Getting started with this Colab notebook is straightforward. Just follow these steps:

1.  **Organize Your Files:** Ensure all necessary files – `req.txt`, your `INPUTPDF` (the PDF you want to process), `Dataformodel.txt`, and `main.py` – are placed in the **same directory** within your Colab environment or local project folder.

2.  **Install Dependencies:** Open your terminal or Colab notebook cell and run the following command to install all required libraries:
    ```bash
    pip install -r req.txt
    ```

3.  **Configure API Keys & File Paths:** Before running, you'll need to update the `main()` function within your `main.py` file. **Fill in all placeholder API keys** and **correctly specify the file addresses** for your `INPUTPDF` and `Dataformodel.txt`.

4.  **Just Run it:**
---

#### How it Works: The Pipeline

1.  PDF to String **(`LLMWhispererV2`):** Your `INPUT PDF` is transformed into a clean, **`string`**.
2.  Question Extraction: An `extract_all_questions` cleans the text and extracts **Numbered Questions**, **Table Questions**, and **Multiple Choice Questions (MCQs)** in dictionary.
3.  **RAG Engine (Google GenAI):** It takes `String` for Question & `Dataformodel.txt` for reference. It generates `String with Q & A`.
4. RAG String to `DOCX`: `convert_string_to_docx` function converts RAG Engine output string to `OutPut2025##.docx` file.
4.  DOCX & PDF Generation **(Adobe PDF Services API):** `docxToPdfConverter` class converts `.docx` file to `output_Output2025##.pdf` file.
---

#### Future Updates:
1. USE Multiple LLMS to create dataformodel.txt
2. USE Computer Vision & More Libraries for Existing PDF filling.
3. Better Data Extraction, Data Classification, Separation and Cleaning.
:)
