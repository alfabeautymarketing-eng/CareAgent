# AgentCare Deployment Guide

## Prerequisites
- A VPS (Ubuntu/Debian recommended).
- Python 3.11+.
- `git`.
- Google Cloud Service Account JSON key.

## Initial Setup on Server

1.  **Clone the Repository**
    ```bash
    git clone https://github.com/YOUR_USERNAME/AgentCare.git
    cd AgentCare
    ```

2.  **Environment Setup**
    Create `.env` file based on `.env.example`:
    ```bash
    cp .env.example .env
    nano .env
    ```
    **Crucial**: set `GOOGLE_CREDENTIALS_FILE=config/credentials.json` (or path to your key).

3.  **Upload Credentials**
    Upload your `credentials.json` to the server:
    ```bash
    # From your local machine
    scp path/to/credentials.json user@your-server-ip:~/AgentCare/config/credentials.json
    ```

4.  **Install Dependencies**
    ```bash
    # Using poetry (recommended)
    pip install poetry
    poetry install
    
    # Or using pip directly
    pip install -r requirements.txt
    ```

## Running the Server

### For Testing (Manual)
```bash
python -m src.main
```

### For Production (Systemd)
Create a service file `/etc/systemd/system/agentcare.service`:

```ini
[Unit]
Description=AgentCare Server
After=network.target

[Service]
User=root
WorkingDirectory=/root/AgentCare
ExecStart=/usr/local/bin/poetry run python -m src.main
Restart=always

[Install]
WantedBy=multi-user.target
```

Enable and start:
```bash
sudo systemctl enable agentcare
sudo systemctl start agentcare
```

## Updating the Server (Continuous Deployment)

### Option 1: Manual Git Pull (Simplest)
When you push changes to GitHub:
1.  SSH into your server.
2.  Run:
    ```bash
    cd AgentCare
    git pull origin main
    sudo systemctl restart agentcare
    ```

### Option 2: Automatic (Webhook)
You can set up a GitHub Action or a webhook listener on the server to auto-pull, but Option 1 is safer for now.

## Troubleshooting
- Check logs: `sudo journalctl -u agentcare -f`
- Check status: `sudo systemctl status agentcare`
