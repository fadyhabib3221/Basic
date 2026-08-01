# Flight Ticket Manager — Deploy It On Your Own Website

This folder is a complete, ready-to-build website version of your Flight Ticket
Manager app. It no longer depends on Claude's internal storage — instead it uses
**Firebase** (a free database) so it works exactly the same, but on any website
you choose (your own domain, Vercel, Netlify, GitHub Pages, etc).

I could not do the deployment itself for you, because the next steps need your
own accounts (Firebase, and a hosting provider) — nobody else can create those
for you, they're tied to your email/login. But every step below is short and
free. Total time: about 15–20 minutes, no coding needed.

---

## Part 1 — Create your free database (Firebase)

1. Go to **https://console.firebase.google.com** and sign in with any Google account.
2. Click **"Add project"** → give it any name (e.g. `flight-tickets`) → keep clicking
   **Continue** with the default options → **Create project**.
3. In the left sidebar, click **Build → Firestore Database** → **Create database**.
   - Choose any location close to you.
   - Choose **"Start in test mode"** (you can tighten security later — see the
     rules comment inside `src/firebaseConfig.js`).
4. In the left sidebar, click the **gear icon → Project settings**.
5. Scroll down to **"Your apps"** → click the **`</>`** (web) icon → give the app
   any nickname → click **Register app**.
6. Firebase will show you a code block that looks like this:
   ```js
   const firebaseConfig = {
     apiKey: "AIza...",
     authDomain: "flight-tickets-xxxxx.firebaseapp.com",
     projectId: "flight-tickets-xxxxx",
     storageBucket: "flight-tickets-xxxxx.appspot.com",
     messagingSenderId: "...",
     appId: "...",
   };
   ```
7. Open the file **`src/firebaseConfig.js`** in this project and replace the
   placeholder values with the ones Firebase just gave you. Save the file.

That's it for the database — it's ready to store your tickets, employees, and
suggestions exactly like before.

---

## Part 2 — Put the code on GitHub (recommended, makes hosting a 2-click job)

1. Go to **https://github.com** and sign in (or create a free account).
2. Click **"New repository"** → name it anything (e.g. `flight-tickets`) →
   keep it **Public** or **Private**, both are fine → **Create repository**.
3. On the new repo page, click **"uploading an existing file"** and drag in
   every file/folder from this project (keep the folder structure — `src/`
   should stay a folder).
4. Click **Commit changes**.

*(If you're comfortable with git commands instead, the usual `git init`,
`git add .`, `git commit`, `git remote add origin ...`, `git push` works too.)*

---

## Part 3 — Deploy it live (free hosting)

### Option A: Vercel (recommended, easiest)
1. Go to **https://vercel.com** → sign up using your GitHub account.
2. Click **"Add New" → "Project"**.
3. Select the `flight-tickets` repository you just created.
4. Vercel auto-detects it's a Vite project — just click **Deploy**.
5. Wait ~1 minute. Vercel gives you a live URL like
   `https://flight-tickets-yourname.vercel.app` — that's your website, live and
   working, with a real database behind it.
6. (Optional) In the Vercel project → **Settings → Domains**, you can attach
   your own custom domain (e.g. `tickets.yourcompany.com`) if you own one.

### Option B: Netlify (just as easy)
1. Go to **https://netlify.com** → sign up using your GitHub account.
2. Click **"Add new site" → "Import an existing project"** → pick your repo.
3. Build command: `npm run build` — Publish directory: `dist` (Netlify usually
   fills these in automatically for a Vite project).
4. Click **Deploy site**. You'll get a live URL in about a minute.

Either option works the same way — pick whichever site you prefer.

---

## Part 4 — First use

1. Open your new live URL.
2. Exactly like before, since no accounts exist yet, you'll see the
   **"Create the admin account"** screen — fill it in once. This will never
   show again afterward (protected by the one-time setup flag, same as before).
3. Sign in, add employees, add tickets — everything works the same as the
   version you used inside Claude, just now hosted on your own website with
   real permanent storage.

---

## Notes

- **Passwords**: like before, this app's login is a simple access gate, not a
  bank-grade authentication system. Don't reuse important passwords for it.
- **Costs**: Firebase's free tier (Spark plan) and Vercel/Netlify's free tiers
  are more than enough for this kind of internal tool — you won't be charged
  unless usage is extremely high.
- **Updating the app later**: any time you want to change something, edit the
  files, push to GitHub (or re-upload), and Vercel/Netlify will automatically
  redeploy the new version within a minute or two.
- If you'd rather use a different database (Supabase, or your own backend)
  instead of Firebase, only `src/storage.js` and `src/firebaseConfig.js` would
  need to change — the rest of the app (`src/App.jsx`) doesn't need to change
  at all.
