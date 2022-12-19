# Corporate Microsoft&trade; Notifier

The following script allows to get PushBullet&trade; notification from corporate version of Microsoft&trade; Teams&reg; and Microsoft&trade; Outlook&reg;.

## Requirements

install the requirements with:

```
pip install -r requirements.txt
```

Download the relevant `geckodriver` from:

https://github.com/mozilla/geckodriver/releases

and put it in this folder.

Get a valid access token from your PushBullet&trade; account:

https://docs.pushbullet.com/

and put it inside a file called `pb.key`.

## Usage

```
python notifier.py
```

insert your corporate email account and password