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
   git clone https://github.com/YOUR_USERNAME/helpdesk-ticket-analyzer.git

2. Install dependencies
   pip install -r requirements.txt

3. Create your .env file
   cp .env.example .env
   Then fill in your real API keys

4. Run
   python main.py