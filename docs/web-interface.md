# Browser interface

BidAgent includes a local Streamlit page for uploading the BOM and tender files, entering bid details and instructions, and downloading the proposal and review report.

Install the project and start the page from the repository root:

```powershell
python -m pip install -e .
python -m streamlit run app.py
```

Streamlit opens the local page in your browser, usually at `http://localhost:8501`. On the first run, upload a finished bid `.docx` in the template field to create the Word template and initial product library. Then upload an `.xlsx` or `.xlsm` BOM, optionally upload tender documents, fill in the project details, and select **Generate technical proposal**. When generation finishes, download the `.docx` and review `.md` files from the assistant response.

The page is intended for local use. Uploaded copies are held temporarily under the ignored `work/uploads/` directory while a bid is being processed; generated files are saved under `output/`. The existing command-line interface remains available.
