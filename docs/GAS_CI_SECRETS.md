# GAS & CI: clasp credentials and GitHub secrets

Короткая инструкция по настройке автоматического деплоя Google Apps Script (GAS) через `clasp` и GitHub Actions.

## Требуемые секреты (рекомендуемые имена)

- `CLASP_CREDENTIALS` — содержимое `~/.clasprc.json` (если вы используете OAuth локального пользователя)
- `CLASP_SERVICE_ACCOUNT` — JSON key сервисного аккаунта (альтернатива для без‑парольного доступа: `clasp login --creds ./sa.json`)
- `SSH_PRIVATE_KEY` — приватный SSH ключ для доступа к серверу (используется `deploy.sh`)
- `SSH_USER` / `SSH_HOST` (опционально) — значения, если `deploy.sh` ожидает конкретные переменные

> Замечание: храните только зашифрованные секреты в GitHub Secrets — никогда не коммитите `~/.clasprc.json` или `sa.json` в репозиторий.

## Как получить `~/.clasprc.json` (локальный OAuth)

1. Установите `clasp` (мы рекомендуем использовать `nvm`):
   ```bash
   npm i -g @google/clasp
   # или: nvm install --lts && npm i -g @google/clasp
   ```
2. Авторизуйтесь:
   ```bash
   clasp login
   ```
3. Содержимое `~/.clasprc.json` можно показать командой:
   ```bash
   cat ~/.clasprc.json
   ```
4. Скопируйте JSON в GitHub secret `CLASP_CREDENTIALS`.

## Сервисный аккаунт (рекомендация для CI)

1. В Google Cloud Console создайте Service Account и скачайте ключ `credentials.json`.
2. Дайте письму service account доступ к необходимым Google Sheets (Share → Editor).
3. В workflow используйте:
   ```bash
   echo "$CLASP_SERVICE_ACCOUNT" > ./sa.json
   clasp login --creds ./sa.json
   ```
4. Поместите `credentials.json` в `CLASP_SERVICE_ACCOUNT` (содержимое JSON) в GitHub secrets.

## Пример: как экспортировать `~/.clasprc.json` в секрет

```bash
# В терминале (локально):
cat ~/.clasprc.json | pbcopy   # macOS: копирует в буфер
# Вставьте значение в Settings → Secrets → New repository secret → Name: CLASP_CREDENTIALS
```

## Проверки и права

- `clasp push -f` из каталога с `appsscript.json` загружает файлы в указанный проект (Script ID в `.clasp.json`).
- Для сервисного аккаунта: убедитесь, что в `appsscript.json` и проекте GAS права на редактирование есть у service account.
- Если в CI используется `clasp login --creds`, убедитесь, что формат JSON корректный и не обрезан при копировании (use raw/plain text).

## Безопасность

- Ограничьте доступ к секретам только админам репозитория.
- Ротация ключей: периодически заново сгенерируйте и обновите `CLASP_SERVICE_ACCOUNT`.
- Для дополнительных мер безопасности используйте GitHub Actions environments с required reviewers.

## Troubleshooting

- `Error: Project settings not found.` — создайте `.clasp.json` в папке с GAS: `printf '{"scriptId":"<ID>","rootDir":"."}' > .clasp.json`
- `clasp login --creds` падает — проверьте, что JSON имеет правильные поля и service account имеет доступ к API (Sheets, Drive, Script)
- Если push работает локально, но падает в CI — убедитесь, что переменная окружения в Workflow содержит полный JSON (проверьте вывод `echo "$CLASP_SERVICE_ACCOUNT" | jq .` локально)

---

Если хотите, могу добавить в репозиторий шаблон `docs/SECRETS_CHECKLIST.md` и пример `GAS_README.md` для поддержки вашей команды.
