# Elevated Frames

Marketing website for **Elevated Frames** — professional drone photography and
cinematic video based in Parry Sound, Ontario.

The entire site is a single self-contained file, `index.html`, styled with
[Tailwind CSS](https://tailwindcss.com/) (loaded from a CDN). There is no build
step — open `index.html` in a browser to preview, and any push to the deployed
branch publishes the changes.

## Project structure

```
index.html        The whole site (markup, styles, and scripts)
images/           Photos used in the gallery and hero
README.md         This file
```

## Adding a video

Videos live in the **Videos** section of the site and are hosted on YouTube.
Adding one is a copy-and-paste — no coding required.

### 1. Upload to YouTube and get the video ID

Upload your video to YouTube, then copy its **video ID** from the URL. The ID is
the part after `v=`:

```
https://www.youtube.com/watch?v=dQw4w9WgXcQ
                                 ^^^^^^^^^^^
                                 this is the ID
```

(For a share link like `https://youtu.be/dQw4w9WgXcQ`, the ID is the part after
the slash.)

### 2. Add a video card to index.html

Open `index.html` and find the videos grid:

```html
<div class="grid md:grid-cols-2 gap-8 max-w-4xl mx-auto" id="videos-grid">
```

Paste this block inside it, replacing `YOUR_VIDEO_ID` and `Your video title`:

```html
<div class="video-card group relative rounded-2xl overflow-hidden bg-zinc-900 border border-zinc-800 aspect-video cursor-pointer shadow-xl"
     data-youtube-id="YOUR_VIDEO_ID"
     data-title="Your video title"></div>
```

That's it. The site automatically builds a clickable thumbnail with a play
button. The YouTube player only loads when a visitor clicks play, which keeps
the page fast.

### 3. Remove the placeholder (first video only)

Until the first real video is added, the section shows a "coming soon" card.
Delete this block once your first video is live:

```html
<div id="video-placeholder" class="md:col-span-2 ...">
  ...
</div>
```

To add more videos later, just paste another `video-card` block — no need to
touch anything else.

### Tips

- **Titles** appear in a caption across the bottom of the thumbnail.
- The thumbnail image is pulled automatically from YouTube; you don't need to
  add one.
- Embeds use `youtube-nocookie.com`, the privacy-friendly YouTube domain.
- A copy of these instructions and the card template also lives in a comment
  right next to the videos grid in `index.html`.

## Editing photos

Gallery photos are `<img>` tags in the **Our Work** section of `index.html`.
Each image has a `data-category` (`real-estate`, `construction`, `events`, or
`landscapes`) that controls which filter button shows it. Add the image file to
the `images/` folder and add a matching `<img>` tag to the gallery.
