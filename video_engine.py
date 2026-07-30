"""
Samsung Creative — animated loop generator.

Takes the exact same layers engine.py already renders (background, product,
logo, headline, features, price, volume badge) and animates them in with a
staggered fade + subtle product zoom, then encodes a short looping MP4 per
format via ffmpeg. Re-run after engine.py — it reuses its render_layers().
"""

import os
import shutil
import subprocess
from PIL import Image
from engine import (
    load_main_model, generate_brief, render_layers, FORMATS, VIDEO_FORMAT, get_format_spec, BASE, OUT_DIR
)

FPS = 30
DURATION_S = 3.5
TOTAL_FRAMES = int(FPS * DURATION_S)

# (layer_name, start_fraction, end_fraction) — alpha ramps 0->1 across that window,
# then holds fully visible for the rest of the loop.
TIMELINE = [
    ("background", 0.00, 0.05),
    ("product",    0.05, 0.28),
    ("logo",       0.22, 0.35),
    ("headline",   0.32, 0.50),
    ("features",   0.48, 0.60),
    ("price",      0.55, 0.75),
    ("volume_badge", 0.70, 0.82),
]
LAYER_ORDER = ["background", "product", "logo", "headline", "features", "price", "volume_badge"]


def ease_out(t):
    return 1 - (1 - t) ** 2


def alpha_at(frame_idx, start_frac, end_frac):
    t = frame_idx / (TOTAL_FRAMES - 1)
    if t <= start_frac:
        return 0.0
    if t >= end_frac:
        return 1.0
    local = (t - start_frac) / (end_frac - start_frac)
    return ease_out(local)


def scale_for_product(frame_idx):
    """Slow continuous zoom for a bit of life, 1.00 -> 1.04 over the whole loop."""
    t = frame_idx / (TOTAL_FRAMES - 1)
    return 1.0 + 0.04 * t


def render_video(format_key, video_dir):
    model = load_main_model(os.path.join(BASE, "samsung-creative-brief-template.xlsx"))
    brief = generate_brief(model)
    layers, meta = render_layers(model, brief, format_key)
    spec = get_format_spec(format_key)
    W, H = spec["size"]

    frames_dir = os.path.join(video_dir, f"_frames_{format_key}")
    os.makedirs(frames_dir, exist_ok=True)

    timeline_map = {name: (s, e) for name, s, e in TIMELINE}

    for i in range(TOTAL_FRAMES):
        frame = Image.new("RGBA", (W, H), (255, 255, 255, 255))
        for name in LAYER_ORDER:
            layer = layers[name]
            a = alpha_at(i, *timeline_map[name])
            if a <= 0:
                continue
            if name == "product" and a > 0:
                # subtle zoom, anchored at layer's own center of mass (approx: canvas center)
                scale = scale_for_product(i)
                new_size = (int(W * scale), int(H * scale))
                scaled = layer.resize(new_size, Image.LANCZOS)
                offset = (-(new_size[0] - W) // 2, -(new_size[1] - H) // 2)
                scaled_canvas = Image.new("RGBA", (W, H), (0, 0, 0, 0))
                scaled_canvas.paste(scaled, offset, scaled)
                layer_img = scaled_canvas
            else:
                layer_img = layer
            if a < 1.0:
                alpha_ch = layer_img.getchannel("A").point(lambda p, a=a: int(p * a))
                layer_img = layer_img.copy()
                layer_img.putalpha(alpha_ch)
            frame = Image.alpha_composite(frame, layer_img)
        frame.convert("RGB").save(os.path.join(frames_dir, f"f_{i:04d}.jpg"), quality=90)

    out_path = os.path.join(video_dir, f"{spec['label'].replace('/', '-')}.mp4")
    subprocess.run([
        "ffmpeg", "-y", "-framerate", str(FPS),
        "-i", os.path.join(frames_dir, "f_%04d.jpg"),
        "-vf", "format=yuv420p",
        "-movflags", "+faststart",
        out_path,
    ], check=True, capture_output=True)

    shutil.rmtree(frames_dir)
    return out_path


def main():
    video_dir = os.path.join(OUT_DIR, "video")
    os.makedirs(video_dir, exist_ok=True)
    path = render_video("video", video_dir)
    size_kb = os.path.getsize(path) / 1024
    print(f"Video (1920×1080): {path} ({size_kb:.0f} KB)")


if __name__ == "__main__":
    main()
