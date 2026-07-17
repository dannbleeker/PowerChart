# Web render repros

Minimal, self-contained repros for PowerPoint-on-the-web rendering bugs, kept here
so they can be attached to [OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues)
issues and re-run without the whole add-in.

## `ellipse-web-repro.yaml` — ellipse vs rectangle fill on web

**Symptom it isolates:** PowerChart's bubble and scatter points (drawn as
`GeometricShapeType.ellipse` + `fill.setSolidColor(...)`) do not appear on
PowerPoint on the web, while rectangles drawn the identical way (chart bars) do.

**Run it:**

1. Install the **Script Lab** add-in (Insert → Get Add-ins → search "Script Lab").
2. Script Lab → **Import** → **paste the YAML** from `ellipse-web-repro.yaml`
   (or import from its raw GitHub URL).
3. Open a **blank slide**, click **Draw rect + 3 ellipses**.
4. Then click **Count shapes on this slide** and read the console at the bottom.

**Read the result:**

| What you see | Meaning |
| --- | --- |
| Rect fills, ellipses don't — but **Count lists all of them** | Created but **not painted** → the web repaint bug (office-js#2699). Zoom in/out or click away+back should make them appear. |
| Rect fills, ellipses don't — **and zoom makes them appear** | Confirms the repaint bug. The shapes are fine; PowerPoint web just didn't repaint. |
| Rect fills, ellipses **never** appear even zoomed | Genuine **ellipse render bug** — file it against office-js with this repro. |
| Ellipse **#3 (sync-first)** fills but **#2 (as-is)** doesn't | The fill must be set in a **separate sync** from the add on web — a one-line fix in the renderer. |
| Ellipse **#4 (no outline)** fills but **#2** doesn't | The white outline is masking a fill that never applied. |

The **quickest** check without Script Lab: in PowerChart, pick the **Bubble** chart
and **Insert into slide** onto the current (visible) slide. If the bubbles render
there but not on the demo deck's appended slides, it's the off-screen repaint bug,
not the ellipse geometry.
