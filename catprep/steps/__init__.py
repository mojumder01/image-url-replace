"""Pipeline steps. Each module has run(ctx) -> dict summary."""

from . import s1_urls, s2_export, s3_trim, s4_images, s5_clean, s5_split, s6_merge, s7_upload, s8_submit

# name -> (module, short description)
STEPS = {
    "urls": (s1_urls, "Add base URL to image keys"),
    "export": (s2_export, "Export seller products from seller-admin + combine"),
    "trim": (s3_trim, "Keep the columns needed for cleaning"),
    "images": (s4_images, "Put new image URLs into the HTML"),
    "clean": (s5_clean, "Clean the HTML (Python rules)"),
    "split": (s5_split, "Split into parts for Claude"),
    "merge": (s6_merge, "Merge Claude's cleaned files"),
    "upload": (s7_upload, "Build upload files + approval file"),
    "submit": (s8_submit, "Upload the files to seller-admin"),
}

# Fully automatic route (no person needed).
AUTO = ["urls", "export", "trim", "images", "clean", "upload", "submit"]

# Manual Claude route: prepare, clean in Claude by hand, then finish.
PREPARE = ["urls", "export", "trim", "images", "split"]
FINISH = ["merge", "upload", "submit"]


def steps_for(mode):
    return AUTO if mode == "auto" else PREPARE + FINISH
