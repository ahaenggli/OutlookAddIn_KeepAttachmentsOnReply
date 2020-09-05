# OutlookAddIn: Keep attachments on reply [![GitHub release (latest by date)](https://img.shields.io/github/v/release/ahaenggli/OutlookAddIn_KeepAttachmentsOnReply?style=social)](https://github.com/ahaenggli/OutlookAddIn_KeepAttachmentsOnReply)
[![paypal](https://www.paypalobjects.com/de_DE/CH/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=S2F6JC7DGR548&source=url)
<a href="https://www.buymeacoffee.com/ahaenggli" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-orange.png" alt="Buy Me A Coffee" height="50px" width="217px" ></a>

## Features
- Bei "antworten" und "allen antworten" werden die Mailanhänge automatisch eingefügt
- Button um bei bestehenden Mails die letzten Anhänge wiedereinzufügen  
    ![FehlendeErmitteln](img/screenshot_button.png)  
    - Bei signierten Mails wird das Mail dafür ohne Zertifikat dupliziert
    - Es können auch mehrere Mails markiert werden 
- Auto-Update Funktion (in Einstellungen deaktivierbar)  
    ![Einstellungen](img/screenshot_settings.png)

## Changelog
... [findet sich hier](CHANGELOG.md) ...

## Auto-Update?
Beim Starten von Outlook, 1x pro 24h, wird auf GitHub die Version überprüft. Gibt es eine neuere Version, wird diese als zip-Datei heruntergeladen und entpackt. Beim nächsten Outlook-Neustart wird die neuere Version dann via ClickOnce-Update nachgeführt.

Wenn Outlook nie geschlossen und gestartet wird, wird auch kein Update installiert.

## Fehler melden
Es dürfen gerne [hier](https://github.com/ahaenggli/OutlookAddIn_KeepAttachmentsOnReply/issues) in GitHub Issues erfasst werden.