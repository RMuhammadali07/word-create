# VPS deployment

1. Install Python 3.11+.
2. Copy `.env.example` to `.env` and set `BOT_TOKEN`, `ADMIN_IDS`, `ADMIN_REVIEW_GROUP_ID`, `PAYMENT_CARD_NUMBER`, and `DATA_DIR`. The bot loads `.env`, not `.env.example`.
3. Install dependencies: `python -m pip install -r requirements.txt`.
4. Start with one worker: `python bot.py`.
5. Keep `data/private` and `data/bot.sqlite3` restricted to the bot OS account.
6. Back up `data/bot.sqlite3` and `data/private` regularly.
7. Add the bot as a Telegram administrator in the review group so it can verify group membership for receipt decisions.
8. Run tests: `python -m pytest -q`.

Telegram group ID should be the numeric ID of the dedicated review group. Run only one polling process at a time.

## User commands

- `/royxatga_olish` checks how many Word generations remain before starting. If none remain, the bot displays the payment card and asks for a receipt. Users upload their own `.docx` template unless the admin enables the shared template in `/admin_jrgrumua`. By default, the shared template is `data/private/templates/6D.docx`; set `SHABLON` to override its location.
- `/sotib_olish` displays the card and receipt instructions. The reviewer enters the amount shown on the receipt; the bot grants one Word generation per 5,000 UZS. The payment must be a whole multiple of 5,000 UZS.
- `/profil` shows remaining Word generations and free late-student additions.
- `/promo CODE` redeems an administrator-created code. The admin creates codes with `/promo_yarat CODE WORD_COUNT MAX_USES`; each user can redeem a code once.
- `/qoshimcha` asks the user to select one of their generated Word files, then adds late students to that file. The first three late-student additions total per user are free; later updates use one Word generation.
- A Word generation is counted only after successful generation.
- The main menu has buttons for roster creation, files, profile, payment, promo codes, late additions, and cancellation.

## Admin user management

- In `/admin_jrgrumua`, open **Foydalanuvchilar** and select a user to see their statistics and actions.
- **Word berish** accepts a positive whole number and adds that many Word generations to the selected user.
- **Word hisobot** and **PDF hisobot** generate a per-user report with statistics, saved text, student details, photos, and receipt images. Original templates, generated Word files, and source media are also sent in a ZIP archive.
- **Ma'lumotlarni tozalash** asks for confirmation, removes the selected user's profile details, roster data, and private files, and anonymizes their account. Payment, promo-code, and admin-grant audit history is retained.
