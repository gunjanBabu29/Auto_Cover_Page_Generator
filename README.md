# Proposal Cover Page Generator

## Run locally

1. Install Python 3.10+.
2. Open a terminal in this folder.
3. Run:

```bash
pip install -r requirements.txt
streamlit run app.py
```

4. Open the local Streamlit URL shown in the terminal.

## What it does

- Select EOI / RFP / Technical Proposal / Financial Proposal.
- Enter project name, reference number and date.
- Upload client/government and funding/bank logos.
- Upload up to 4 project images.
- Add 1–5 submitting organizations with roles and logos.
- Preview the cover.
- Export a high-resolution PNG and a Word DOCX.

The current design is based on the uploaded EOI cover-page visual structure: top logos, central document heading, rounded project-information panel, date capsule, project-image collage and Submitted By panel.

## Important

The first version exports the completed design as one high-resolution image inside Word. This is intentional because it keeps the complex cover layout stable. The next upgrade can make each text/logo/image element independently editable in Word.
