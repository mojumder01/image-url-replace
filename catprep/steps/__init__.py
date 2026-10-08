"""Pipeline steps in order. Each module has run(ctx) -> dict summary."""

from . import s1_urls, s2_export, s3_trim, s4_images, s5_split, s6_merge, s7_upload

# name -> (module, short description)
STEPS = {
    "urls": (s1_urls, "Add base URL to image keys"),
    "export": (s2_export, "Export seller products from seller-admin + combine"),
    "trim": (s3_trim, "Keep the columns Claude needs"),
    "images": (s4_images, "Put new image URLs into the HTML"),
    "split": (s5_split, "Split into parts for Claude"),
    "merge": (s6_merge, "Merge Claude's cleaned files"),
    "upload": (s7_upload, "Build upload files + approval file"),
}

PREPARE = ["urls", "export", "trim", "images", "split"]
FINISH = ["merge", "upload"]
