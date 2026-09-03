# Putting Gridlint online

The judges need a URL they can open. Two free options, both about ten minutes.
Neither needs a credit card.

Whichever you choose, **check it from your phone afterwards** — a live demo that only works
on your own laptop is worse than no live demo.

---

## Option A — Hugging Face Spaces (recommended)

Spaces runs the `Dockerfile` in this repository unchanged, and it stays up.

1. Sign in at <https://huggingface.co> and choose **New Space**.
2. Owner: your account. Space name: `gridlint`. **SDK: Docker** (not Gradio — the app is
   FastAPI and the Gradio template will not run it). Hardware: **CPU basic (free)**.
   Visibility: **Public**.
3. Push this repository into the Space:

   ```bash
   git remote add space https://huggingface.co/spaces/<your-username>/gridlint
   git push space main
   ```

   If it asks for a password, use an access token from
   <https://huggingface.co/settings/tokens> with **write** permission.
4. Wait for the build (three to five minutes). The URL is
   `https://<your-username>-gridlint.hf.space`.

Things worth knowing:

- The container listens on **port 7860**, which the `Dockerfile` already does.
- A Space **sleeps after 48 hours of inactivity** and the first request afterwards takes
  30–60 seconds. Judging runs 16–20 September, so **open your own Space once on the 15th**
  and it will be warm.
- The free tier gives you no persistent disk you can rely on. `GRIDLINT_DATA=/data` is
  writable, so sign-up and uploads work, but treat the database as disposable. That is fine
  for a demo: `/api/demo` needs no account at all.

---

## Option B — Render

1. Sign in at <https://render.com>, **New → Web Service**, connect the GitHub repository.
2. Runtime **Docker**. Instance type **Free**. Region: whichever is nearest your judges
   (Oregon is a safe default).
3. Deploy. The URL is `https://gridlint-xxxx.onrender.com`.

Things worth knowing:

- A free service **spins down after 15 minutes of inactivity** and takes about a minute to
  wake. That is a bad first impression for a judge who clicks your link cold.
- Mitigate it: put the Space or Render URL into a free uptime pinger
  (<https://uptimerobot.com>, 5-minute interval) for the judging week, or use Option A.

---

## Before you paste the link into Devpost

Run through this on the deployed site, not on localhost:

- [ ] The landing page loads and the three proof numbers are visible.
- [ ] **Check the example model** produces the report, with runway 38.6 → 5.2.
- [ ] **Explain in plain English** writes the notes (they replay from the committed
      fixtures, so this works with no API key).
- [ ] Uploading `samples/clean-model.xlsx` reports **no defects** — showing it stays quiet
      is as convincing as showing it finds things.
- [ ] The grid renders and **show formulas** works.
- [ ] It looks right on a phone.
- [ ] `/<your-url>/api/health` returns `{"ok": true, ...}`.

## Optional: turn on live explanations

Everything works without this. If you want the notes generated live rather than replayed,
add one environment variable in the Space or Render dashboard:

```
ANTHROPIC_API_KEY=...     # or OPENAI_API_KEY / GROQ_API_KEY / GEMINI_API_KEY
```

Groq and Google both have a free tier that is more than enough for a demo. Note that this
is an **API key**, which is billed separately from any ChatGPT or Claude subscription — a
subscription does not give you one.

Only finding metadata is ever sent to the model. The workbook itself never leaves the
server.
