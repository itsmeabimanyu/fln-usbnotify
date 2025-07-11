## Requirements

```bash
pip install psutil python-dotenv schedule
```

Additional Dependencies

- for Linux 

    ```bash
    pip install pyudev
    ```
- for Windows 

    ```bash
    pip install pywin32 wmi
    ```

Create a file named `.env` in your project directory with the following content:

```bash
EMAIL_HOST=
EMAIL_PORT=
EMAIL_HOST_USER=
EMAIL_HOST_PASSWORD=
TO_EMAIL=
VALID_IP_PREFIX=
```