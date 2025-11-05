An automated Python pipeline that:

Watches for a .docx file update,

Converts it to clean HTML,

Pushes the updated HTML to GitHub Pages repository.

Here’s a complete, working setup outline — including Python script + GitHub Actions YAML + repository structure.

🧱 Folder Structure
docx-to-html-site/
│
├── convert_and_push.py
├── requirements.txt
├── input/
│   └── song_notations.docx
├── output/
│   └── song_notations.html
└── .github/
    └── workflows/
        └── convert_and_deploy.yml