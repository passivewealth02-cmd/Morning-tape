# Run Design Autopilot Coach on your own computer

No Vercel, no cloud. Everything runs locally. The only external call is to
OpenAI when you press **Generate image** — that goes straight from your
machine to OpenAI with your own key.

You need to do this once. After that, it's a single command to start.

## 1. Install Node.js (one time)

Download the **LTS** installer from <https://nodejs.org> and run it. This
gives you `node` and `npm`. To confirm, open a terminal (Terminal on Mac,
PowerShell on Windows) and run:

```bash
node --version
```

You should see something like `v20.x` or newer.

## 2. Get the code

If you have `git`:

```bash
git clone https://github.com/passivewealth02-cmd/Morning-tape.git
cd Morning-tape
git checkout claude/design-autopilot-coach-w8k6jm
```

No git? On the GitHub page, switch to the branch
`claude/design-autopilot-coach-w8k6jm`, click **Code → Download ZIP**, unzip
it, and `cd` into the folder in your terminal.

## 3. Install the dependencies (one time)

This project uses **pnpm**. Install it once, then install the packages:

```bash
npm install -g pnpm
pnpm install
```

## 4. Add your OpenAI key

Create a file named **`.env.local`** in the project folder with one line
(use a freshly created key — never reuse one you've shared anywhere):

```
OPENAI_API_KEY=sk-your-new-key-here
```

That's the only variable the Coach needs. `.env.local` is git-ignored, so it
never leaves your machine.

> The Coach still runs without this — you just get a "add your key" message on
> the image button instead of a generated PNG.

## 5. Start it

```bash
pnpm dev
```

Open **<http://localhost:3000/coach>** in your browser. That's the whole app,
running on your computer:

- fill the form (or click an example) → **Generate recipe**
- upload up to 10 **reference images** to steer the style
- **Generate image** → a transparent PNG you can download

To stop the app, press `Ctrl + C` in the terminal. To start it again later,
just `cd` into the folder and run `pnpm dev`.

## Notes

- **Cost:** image generation uses your OpenAI credits (~a few cents each with
  `gpt-image-1`). Make sure that account has credit or a payment method.
- **`gpt-image-1` verification:** some new OpenAI accounts must complete a
  quick organization verification before this model is enabled. If the button
  reports "must be verified", do that on the OpenAI dashboard.
- **Port in use?** If `3000` is taken, run `pnpm dev -- -p 3005` and use
  `http://localhost:3005/coach`.
