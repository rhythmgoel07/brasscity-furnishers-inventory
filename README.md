# Furniture Inventory — Shared, Multi-User (Firebase)

You manage inventory from anywhere; your staff view live stock from the shop WiFi (or
anywhere too, technically) — no login needed for them. Real-time: when you change stock,
staff see it update automatically, no refresh needed.

## How this differs from the offline PWA version

- **Data lives in Firestore** (a cloud database), not on one device — that's what makes
  it shared and accessible from anywhere.
- **You sign in** (email + password) to unlock Manage mode. Staff use Browse mode with no
  login at all.
- **Real-time sync** — Firestore pushes updates to every open device instantly. No pull-
  to-refresh, no polling.
- **Needs internet** to sync — this is the direct tradeoff for "accessible from anywhere."
  It still shows your last-synced data if connectivity drops (via Firestore's own offline
  cache), but new changes made elsewhere won't appear until you're back online.
- **Excel import and export are both included**, confirmation-gated. Import *merges*:
  matching SKUs are updated, new ones added, and anything already in the database but
  absent from the file is left alone — an import can never wipe the collection. Rows
  with no photo now keep whatever photo is already stored. **Backup (JSON)** is there
  too as a safety-net download.

---

## Part 1 — Create your Firebase project (~10 minutes, one time)

1. Go to **console.firebase.google.com**, sign in with any Google account
2. Click **Create a project** → name it anything (e.g. `brasscity-inventory`) → you can
   disable Google Analytics for this (not needed) → **Create project**

### Enable Firestore (the database)
3. In the left sidebar, click **Build → Firestore Database** → **Create database**
4. Choose **Start in production mode** → pick a location close to you (any nearby region)
   → **Enable**

### Set the security rules
5. Still in Firestore, click the **Rules** tab at the top
6. Delete everything there and paste in the contents of **`firestore.rules`** (included
   in this download) → click **Publish**

### Enable Authentication (your manager login)
7. Left sidebar → **Build → Authentication** → **Get started**
8. Click **Email/Password** in the provider list → toggle it **Enable** → **Save**
9. Go to the **Users** tab → **Add user** → enter your own email and choose a password →
   **Add user**. This is the account you'll sign in with as manager — there's no
   "sign up" screen in the app itself, you add accounts here directly.

### Mark yourself as a manager
10. **Authentication → Users** → copy the **User UID** shown next to your email.
11. **Firestore Database → Data** → **Start collection** → collection ID `managers` →
    for **Document ID**, paste that UID → add no fields at all → **Save**.

    Being *signed in* is no longer enough to edit anything. Only accounts listed in
    `managers` can write. That's deliberate: with anonymous sign-in switched on, the
    old "any signed-in user can write" rule would have meant anyone at all could edit
    or delete your entire inventory. To give a second person edit access later, repeat
    steps 10–11 with their UID.

### Get your config values
12. Click the **⚙️ gear icon** (top-left, next to "Project Overview") → **Project settings**
13. Scroll to **Your apps** → click the **`</>`** (web) icon → give it any nickname → **Register app**
14. You'll see a code block with `const firebaseConfig = { apiKey: "...", ... }` —
    copy those exact values

## Part 2 — Configure the app

1. Open **`firebase-config.js`** from this download in any text editor (Notepad is fine)
2. Replace each `PASTE_YOUR_...` placeholder with the matching value you just copied
3. Save the file

That's the only file you need to edit.

## Part 3 — Deploy to GitHub Pages

Same process as before:
1. Create a GitHub repository, upload every file in this folder (keeping the folder
   structure — `icons/` included) — **except `firestore.rules`**, which isn't part of
   the app itself, it's just for pasting into the Firebase console (you already did that
   in Part 1)
2. Repo **Settings → Pages** → Source: "Deploy from a branch" → `main` / `/(root)` → Save
3. Your live URL appears after about a minute:
   `https://<your-username>.github.io/<repo-name>/`

## Part 4 — Install on phones

**You (Manager):**
- Open the URL → tap **⚙️ Manage** → sign in with the email/password from Part 1, step 9
- Add to Home Screen the same way as before (Android: ⋮ menu; iPhone: Share → Add to Home Screen)

**Staff:**
- Open the same URL → they land in **🏬 Browse** mode automatically — no login, no setup
- They can search by text, photo, or scan QR tags, but won't see any edit controls
- Add to Home Screen too, if you want it to feel like an app for them as well

Everyone opening that same URL sees the same shared, live inventory.

## A note on the manager password

Manager accounts are created directly in the Firebase console (Part 1, step 9) — this app
has no "create account" screen.

Giving someone edit access takes **two** steps, not one: create the user under
**Authentication → Users**, *and* add their UID to the `managers` collection (Part 1,
steps 10–11). Creating the account alone lets them sign in but not change anything, which
is the intended behaviour — removing someone's edit rights later is then just a matter of
deleting their document from `managers`.

## Updating the app later

If you change any files and re-upload to GitHub, bump `CACHE_VERSION` at the top of
`sw.js` (e.g. `v1` → `v2`) so phones with the app already installed pick up the change
next time they open it with a connection available.

## Cost

Firebase's free tier ("Spark plan") comfortably covers a single shop's usage — reads/
writes/storage limits are generous for an inventory this size. You'd need very heavy,
sustained multi-store usage before hitting any billing at all, and Firebase won't charge
you unless you explicitly upgrade to a paid plan.


---

## Item Tags

Categories are gone. Items carry **tags** instead, and a tag exists only if you
defined it in **Manage → Item Tags**.

- Tags are a single word, letters and digits only, stored lowercase and shown
  capitalised. That's what lets the Excel column be comma-separated safely.
- On the item form, the tag box **searches** your defined tags. Click a suggestion
  to add it; each one appears as a pill with a × to remove. You cannot create a
  tag from that box — that's the point.
- In Browse, tag chips combine with **AND**: selecting *dining* and *tables* shows
  only items carrying both. Each chip narrows.
- Deleting a tag removes it from every item that uses it. The count is shown
  before you confirm, and it can't be undone.

**Coming from categories:** a *Convert categories to tags* button appears in Manage
while any item still has one. Each category becomes a tag, and multi-word ones like
"Dining Set" split into two tags. Nothing is retyped.

### Tags in Excel

The `Category` column is replaced by `Tags`, holding names separated by commas:

| SKU | Item Name | Tags |
|---|---|---|
| BF-61 | Dining Table | dining, tables |

A second worksheet named **Tags** lists every valid tag. It's regenerated on each
export — editing it does nothing.

Because tags never contain spaces, the import is forgiving: `dining, tables`,
`dining tables` and `Dining;Tables` all parse identically, and case is ignored.

**Unknown tags block the import.** If the file contains a tag you haven't defined —
a typo like `dinning`, say — nothing is imported and you're shown which tags and
which SKUs. Fix the spelling, or define the tag first. In the exported file such
cells are also highlighted red as you type.

---

## Editing the Excel export

Three columns in the exported workbook are **calculated**, and are labelled
`(auto)` with grey italic text:

| Column | Worked out from |
|---|---|
| Margin % (auto) | Cost Price and Selling Price |
| Stock Value (auto) | Cost Price × Stock Qty |
| Stock Status (auto) | Stock Qty vs Reorder Level |

Editing them does nothing — the import ignores those columns and recalculates all
three from the real numbers.

**To put an item back in stock, change its `Stock Qty`, not `Stock Status`.** An item
shows OUT OF STOCK whenever Stock Qty is 0, LOW STOCK when it's at or below the Reorder
Level, and IN STOCK above that.

If you do edit a status by mistake, the import preview now warns you before anything is
written, listing the rows affected.

---

## Part 5 — Google Sheets two-way sync (optional)

Skip this and everything else still works; the Sync button just explains it isn't
configured.

### Create the OAuth client ID
1. Go to **console.cloud.google.com**, and pick the same project as your Firebase app
   (Firebase projects *are* Google Cloud projects — it will be in the list).
2. **APIs & Services → Library** → search **Google Sheets API** → **Enable**.
3. **APIs & Services → OAuth consent screen** → User type **External** → fill in an app
   name and your email → **Save and continue**.
4. On the **Scopes** step, add `https://www.googleapis.com/auth/drive.file`. It will be
   listed as **non-sensitive**, which is the whole point: non-sensitive scopes don't
   require Google's verification review or a security assessment. Don't add anything
   broader — `drive.readonly` and `drive` are restricted scopes and drag you into a paid
   annual audit.
5. **Publish** the consent screen. With only non-sensitive scopes there's nothing to
   verify, and leaving it in Testing forces you to re-consent every week.
6. **APIs & Services → Credentials → Create credentials → OAuth client ID** → type
   **Web application**. Under **Authorised JavaScript origins** add your live URL
   exactly, with no trailing slash:
   `https://<your-username>.github.io`
7. Copy the client ID and paste it into `googleClientId` in `firebase-config.js`.

### First run
Open the app as manager → **🔄 Sync with Sheet** → consent once. The app creates a
spreadsheet in your Drive called *Furniture Inventory (synced)* and fills it from your
current inventory. That file is the one it will keep using.

### How the sync actually behaves

It is **not** a live mirror. Nothing moves until you press Sync, and nothing is written
until you've seen the plan and pressed Apply.

- Merging happens per **field**, not per row. Change the price in the sheet and the stock
  in the app on the same item, and both apply — that isn't a conflict.
- If only one side changed a field, that side wins. The app keeps a **baseline** of what
  both sides looked like after the last successful sync, which is how it can tell a real
  edit from a value nobody touched. Without that, a stale sheet cell would happily
  overwrite a real edit.
- If **both** sides changed the same field to different values, it stops and asks. You see
  the old value, the app's value and the sheet's value, and pick. It never guesses.
- A row **missing** from the sheet is never deleted automatically. Rows disappear from
  spreadsheets far too easily. Those are listed with tick boxes, and unticked items are
  simply kept.
- Column **`_id`** is hidden and machine-managed. Don't delete or edit it — it's what lets
  you rename a SKU in the sheet without the app reading it as "one item deleted, a
  different one added" and losing the attached photo.
- Columns are matched by **header name**, so inserting a column mid-sheet is safe.
  Renaming or deleting a header aborts the sync before anything is read or compared.
- The baseline is written **last**. If a sync fails halfway, your sync state isn't
  advanced and running it again simply retries the same work.

### The one thing to be careful about

Conflicts are a judgement call made from three numbers on a screen. Whichever value you
don't pick is gone. They should be rare — they only happen when you've edited the same
field in both places between syncs — but don't click through them on autopilot.

The safest habit is to **sync before you start editing the sheet, and again when you
finish**. That keeps the two sides close together and conflicts near zero.

---

## Photos

Photos no longer live inside the item records. They're in a separate `itemPhotos`
collection, fetched only when a card scrolls into view. For a 500-item catalogue that's
roughly 300 KB on opening the app instead of about 10 MB, and adjusting stock no longer
re-uploads the photo along with it.

Because the weight is gone, photos are now stored at two sizes: a small one for the grid
and a 900px one shown when you tap a product. Tap any card to enlarge.

If you have items from the older version, a **📦 Move photos out of items** button appears
in Manage mode. It writes the new photo record *before* clearing the old field, so an
interruption leaves duplicated data rather than missing data, and re-running it is safe.

Photos deliberately stay in Firestore rather than Cloud Storage: since February 2026
Cloud Storage requires the Blaze plan and a linked card, and at this catalogue size the
only thing it would add is CDN caching.
