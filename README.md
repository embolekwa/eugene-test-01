# The Science & Craft of Ministering the Word

A Kingdom Restoration Church training deck on preparing and delivering the word.

| File | What it is |
|---|---|
| `Ministering-the-Word.pptx` | The presentation (16:9, 10 slides) |
| `build_deck.js` | Generator — edit this and rebuild rather than editing the .pptx by hand |
| `assets/` | Logo derivatives (transparent PNGs) cut from the supplied JPEG |

## Slides

1. Title
2. The craft — what listeners are reaching for
3. Scripture — 2 Timothy 2:15 (NKJV)
4. Clarity
5. Flow — introduction, body, conclusion
6. Encouragement — with 1 Corinthians 14:3 (NIV)
7. Balance your biases — Old and New Testament
8. The anointing you carry
9. Preparing Sunday's sermon — purpose informs content
10. Close

Every slide carries speaker notes.

## Branding

Colours are taken from the logo: purple `#61116A`, gold `#AE9600`, with a deep
plum `#330A3A` for the dark slides. The logo appears full-colour on light
slides and in solid white on dark ones; `assets/` also holds black-free
monochrome variants (`logo_mark_purple.png`, `logo_mark_gold.png`) if you want
a different treatment.

## Rebuilding

```bash
npm install pptxgenjs react react-dom react-icons sharp
node build_deck.js
```

Icons are [Feather](https://feathericons.com/) via `react-icons`, rasterised at
build time, so the deck has no external dependencies once generated.
