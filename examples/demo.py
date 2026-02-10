from md2word import fill_template

# 1. Your Data (Python Dictionary)
data = {
    "date": "2026-02-09",
    "start_time": "18:00",
    "end_time": "20:00",
    "pricetable": [
        {"item": "Coffee with milk and honey", "price": "25"},
        {"item": "Waffle", "price": "35"}
    ],
    "cases": [
        {
            "saknr": "10/26",
            "content": "# Annual Meeting on steroids\n\nWelcome to the **annual meeting**.\n\nAgenda:\n1. Approval of accounts\n2. Election of board"
        },
        {
            "saknr": "11/26",
            "content": "## Project X\n\nStatus: *Delayed*.\n\nSee [Jira Board](https://jira.example.com) for details."
        }
    ]
}


# 2. Run it
# (Make sure you have a template.docx in the folder)
fill_template("./template_md2word.docx", data, "./output.docx")
