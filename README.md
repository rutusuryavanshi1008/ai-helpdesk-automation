# 🎫 Helpdesk Ticket Analyzer

An AI-powered helpdesk analytics pipeline that automatically classifies tickets, generates Excel reports with charts, and emails summaries to managers.

![Demo](demo.gif)

## 🚀 Features
- Loads and cleans helpdesk ticket data
- Generates 4 charts (bar, pie, heatmap)
- Creates multi-sheet Excel report
- AI classifies each ticket using Groq (LLaMA)
- Auto-emails the report

## 🛠️ Tech Stack
- Python, Pandas, Matplotlib
- OpenPyXL, Groq API, LLaMA 3.1
- python-dotenv, smtplib

## ⚙️ Setup
1. Clone the repo
2. Run: pip install -r requirements.txt
3. Copy .env.example to .env and fill in your keys
4. Run: python analysis.py