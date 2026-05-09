# ArchLens AI Deployment

## Streamlit Community Cloud
1. Upload `app.py` and `pdf_summary.py` to GitHub.
2. Add `requirements.txt`.
3. Deploy with main file path `app.py`.
4. Set `OPENAI_API_KEY` in secrets.

## Render
Build command:
`pip install -r requirements.txt`

Start command:
`streamlit run app.py --server.port $PORT --server.address 0.0.0.0`

Required environment variables:
- `OPENAI_API_KEY`
- `ARCHLENS_API_URL`
- `ARCHLENS_WEBHOOK_SECRET`
- `ADMIN_EMAILS=salman@sydesignstudio.co.uk,info@sydesignstudio.co.uk`

## Website deployment
Host ArchLens AI on a subdomain such as:
`archlens.sydesignstudio.co.uk`
Then link to it from your SY Design Studio website.
