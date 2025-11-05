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

**Step to** Enable GitHub Pages

Go to Repository → Settings → Pages

Under Build and deployment, choose:

Source: GitHub Actions

GitHub will automatically publish HTML in /output folder.

Your live site link:
https://synapticrumble.github.io/songnotes/song_notations.html

💡 Optional Enhancements

Add custom style.css and link it inside the generated HTML:

html = "<link rel='stylesheet' href='style.css'>" + html

