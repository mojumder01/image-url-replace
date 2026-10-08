"""Step 8: upload the files to seller-admin (update files first, then the approval file).

Every page/button name comes from config.toml [site_upload], so it can be adjusted
without code changes. A screenshot is saved after each upload in 8_upload as proof.
When [site_upload] enabled = false, this step only reminds you to upload by hand.
"""

import os
import time

from .. import site
from ..common import StepError, info, list_xlsx, warn
from ..site import click


def _goto(page, menu_path):
    for text in menu_path:
        click(page.locator(f"text={text}").first, page, f"'{text}'")
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(600)


def _upload_one(page, path, target, timeout_s, shot_path):
    """target: dict with menu, file_input, submit_button, success_text, error_text."""
    _goto(page, target["menu"])
    page.locator(target.get("file_input") or "input[type=file]").first.set_input_files(path)
    page.wait_for_timeout(800)
    if target.get("submit_button"):
        click(page.locator(target["submit_button"]).first, page, "upload button", 5000)
    success = [t.lower() for t in target.get("success_text", [])]
    errors = [t.lower() for t in target.get("error_text", [])]
    end = time.time() + timeout_s
    result = "unknown"
    while time.time() < end:
        page.wait_for_timeout(2000)
        try:
            body = page.locator("body").inner_text(timeout=2000).lower()
        except Exception:
            continue
        if any(t in body for t in errors):
            result = "error"
            break
        if any(t in body for t in success):
            result = "ok"
            break
    page.screenshot(path=shot_path, full_page=True)
    return result


def run(ctx):
    cfg = ctx.cfg.get("site_upload", {})
    out_dir = ctx.run.folder("upload")
    updates = [p for p in list_xlsx(out_dir) if os.path.basename(p).lower().startswith("update_part")]
    approval = ctx.run.file("approval")
    if not cfg.get("enabled"):
        info("Uploading to seller-admin is switched off ([site_upload] enabled = false).")
        info(f"Upload these by hand from {out_dir}:")
        for p in updates + ([approval] if os.path.isfile(approval) else []):
            info(f"    {os.path.basename(p)}")
        return {"uploaded_to_site": False}
    if not updates:
        raise StepError("No update_part files to upload.")

    timeout = int(cfg.get("wait_seconds", 600))
    done = []
    with site.session(ctx.cfg) as page:
        for path in updates:
            name = os.path.basename(path)
            info(f"Uploading {name} ...")
            result = _upload_one(page, path, cfg["update"], timeout,
                                 os.path.join(out_dir, f"site_{os.path.splitext(name)[0]}.png"))
            done.append(f"{name}: {result}")
            if result != "ok":
                raise StepError(f"Upload of {name} did not succeed ({result}). See the screenshot in {out_dir}. "
                                "The approval file was NOT uploaded.")
            info(f"  {name}: uploaded")
        if cfg.get("upload_approval", True) and os.path.isfile(approval):
            info("Uploading Product-approval.xlsx ...")
            result = _upload_one(page, approval, cfg["approval"], timeout,
                                 os.path.join(out_dir, "site_Product-approval.png"))
            done.append(f"Product-approval.xlsx: {result}")
            if result != "ok":
                raise StepError(f"Upload of Product-approval.xlsx did not succeed ({result}).")
        elif not os.path.isfile(approval):
            warn("No Product-approval.xlsx to upload.")
    return {"uploaded_to_site": True, "results": "; ".join(done)}
