#!/usr/bin/env python3
"""
VE3 Simple - Worker tạo ảnh qua server mode

Flow:
1. Load Excel (PromptWorkbook)
2. Tạo reference images (nv/, loc/) qua server
3. Tạo TẤT CẢ scene images (không chia chẵn/lẻ) qua server
4. Tạo thumbnail (nhân vật chính → thumb/)

Usage:
    worker = VE3Worker(project_dir, config, log_func)
    result = worker.run()
"""

import sys
import os
import time
import json
import shutil
from pathlib import Path
from typing import Optional, Callable, Dict, List, Any
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# Đảm bảo import modules từ thư mục ve3
VE3_DIR = Path(__file__).parent
sys.path.insert(0, str(VE3_DIR))

from modules.excel_manager import PromptWorkbook, Character, Scene
from modules.google_flow_api import (
    GoogleFlowAPI, GeneratedImage, ImageInput, ImageInputType,
    AspectRatio, ImageModel, VideoAspectRatio, VideoModel
)
from modules.server_pool import ServerPool


class VE3Worker:
    """Worker tạo ảnh từ Excel qua server mode."""

    def __init__(
        self,
        project_dir: str,
        config: Dict[str, Any],
        log_func: Callable = None,
        progress_func: Callable = None,
        on_item_status: Callable = None
    ):
        """
        Args:
            project_dir: Đường dẫn thư mục project (chứa Excel + output)
            config: Dict cấu hình từ settings.yaml
            log_func: Callback log(msg, level) - hiển thị lên GUI
            progress_func: Callback progress(phase, current, total, detail)
            on_item_status: Callback (item_type, item_id, status, image_path)
        """
        self.project_dir = Path(project_dir)
        self.config = config
        self.log = log_func or (lambda msg, level="INFO": print(f"[{level}] {msg}"))
        self.progress = progress_func or (lambda *a, **kw: None)
        self.on_item_status = on_item_status or (lambda *a, **kw: None)
        self._stop_flag = False

        # Paths
        self.nv_dir = self.project_dir / "nv"
        self.img_dir = self.project_dir / "img"
        self.vid_dir = self.project_dir / "vid"
        self.thumb_dir = self.project_dir / "thumb"

        # Server config
        self.server_url = config.get("local_server_url", "")
        self.server_list = config.get("local_server_list", [])
        self.bearer_token = config.get("flow_bearer_token", "")
        self.flow_project_id = config.get("flow_project_id", "")
        self.timeout = config.get("flow_timeout", 120)
        self.retry_count = config.get("retry_count", 3)

        # Aspect ratio
        ar_str = config.get("flow_aspect_ratio", "landscape").upper()
        self.aspect_ratio = getattr(AspectRatio, ar_str, AspectRatio.LANDSCAPE)

        # Concurrent prompts
        self.max_concurrent = config.get("max_concurrent", 1)

        # Server pool
        self.pool = None
        self._init_server_pool()

    def _init_server_pool(self):
        """Khởi tạo ServerPool từ config."""
        pool_config = {}
        if self.server_list:
            pool_config["local_server_list"] = self.server_list
        elif self.server_url:
            pool_config["local_server_url"] = self.server_url

        if pool_config:
            self.pool = ServerPool(pool_config, log_callback=self.log)
            self.pool.refresh_all()
            self.log(f"Server pool: {len(self.pool.servers)} server(s)")
        else:
            self.log("Không có server URL!", "ERROR")

    def stop(self):
        """Dừng worker."""
        self._stop_flag = True
        self.log("Đang dừng worker...")

    def run(self) -> Dict[str, Any]:
        """
        Pipeline chính: references → scenes → thumbnail.

        Returns:
            Dict kết quả: {success, total, completed, failed, errors}
        """
        result = {"success": False, "total": 0, "completed": 0, "failed": 0, "errors": []}

        if not self.pool:
            result["errors"].append("Không có server URL")
            return result

        # Tìm Excel file
        excel_path = self._find_excel()
        if not excel_path:
            result["errors"].append("Không tìm thấy file Excel trong project")
            return result

        self.log(f"Loading Excel: {excel_path.name}")

        try:
            wb = PromptWorkbook(str(excel_path))
            wb.load_or_create()
        except Exception as e:
            result["errors"].append(f"Lỗi đọc Excel: {e}")
            return result

        # Đọc bearer token từ Excel config nếu chưa có
        if not self.bearer_token:
            self.bearer_token = wb.get_config_value("flow_bearer_token") or ""
        if not self.flow_project_id:
            self.flow_project_id = wb.get_config_value("flow_project_id") or ""

        # Auto-strip prefix "Bearer " nếu user nhập nhầm
        if self.bearer_token.lower().startswith("bearer "):
            self.bearer_token = self.bearer_token[7:].strip()
            self.log("Đã tự động bỏ prefix 'Bearer ' khỏi token", "WARN")

        if not self.bearer_token:
            result["errors"].append("Thiếu bearer token! Nhập trong GUI hoặc sheet config")
            return result

        if not self.bearer_token.startswith("ya29."):
            result["errors"].append(
                f"Bearer token không hợp lệ (phải bắt đầu bằng 'ya29.'). "
                f"Token hiện tại: '{self.bearer_token[:20]}...'. "
                f"Hãy nhập lại token trong GUI (không cần chữ 'Bearer')"
            )
            return result

        # Tạo thư mục output
        self.nv_dir.mkdir(parents=True, exist_ok=True)
        self.img_dir.mkdir(parents=True, exist_ok=True)
        self.vid_dir.mkdir(parents=True, exist_ok=True)

        # === PHASE 1: References ===
        self.log("=" * 50)
        self.log("PHASE 1: Tạo ảnh nhân vật & địa điểm")
        self.log("=" * 50)
        ref_result = self._generate_references(wb)
        if self._stop_flag:
            result["errors"].append("Đã dừng bởi user")
            return result

        # === PHASE 2: Scene Images ===
        self.log("")
        self.log("=" * 50)
        self.log("PHASE 2: Tạo ảnh các cảnh")
        self.log("=" * 50)
        scene_result = self._generate_scenes(wb)
        if self._stop_flag:
            result["errors"].append("Đã dừng bởi user")
            return result

        # === PHASE 3: Thumbnail ===
        self.log("")
        self.log("=" * 50)
        self.log("PHASE 3: Tạo Thumbnail")
        self.log("=" * 50)
        self._generate_thumbnail(wb)

        # === PHASE 4: Videos (Image-to-Video) ===
        self.log("")
        self.log("=" * 50)
        self.log("PHASE 4: Tạo Video từ ảnh")
        self.log("=" * 50)
        vid_result = self._generate_videos(wb)
        if self._stop_flag:
            result["errors"].append("Đã dừng bởi user")
            return result

        # Tổng kết
        img_total = ref_result["total"] + scene_result["total"]
        img_done = ref_result["completed"] + scene_result["completed"]
        img_fail = ref_result["failed"] + scene_result["failed"]
        vid_total = vid_result["total"]
        vid_done = vid_result["completed"]
        vid_fail = vid_result["failed"]

        result["success"] = (img_fail + vid_fail) == 0 and (img_done + vid_done) > 0
        result["total"] = img_total + vid_total
        result["completed"] = img_done + vid_done
        result["failed"] = img_fail + vid_fail

        self.log("")
        self.log("=" * 50)
        status = "HOÀN THÀNH" if result["success"] else "CÓ LỖI"
        self.log(f"KẾT QUẢ: {status} - Ảnh: {img_done}/{img_total}, Video: {vid_done}/{vid_total}")
        self.log("=" * 50)

        return result

    def _find_excel(self) -> Optional[Path]:
        """Tìm file Excel trong project dir."""
        # Tìm *_prompts.xlsx trước
        for f in self.project_dir.glob("*_prompts.xlsx"):
            if not f.name.startswith("~"):
                return f
        # Fallback: bất kỳ .xlsx
        for f in self.project_dir.glob("*.xlsx"):
            if not f.name.startswith("~"):
                return f
        return None

    # =========================================================================
    # PHASE 1: Reference Images
    # =========================================================================

    def _generate_references(self, wb: PromptWorkbook) -> Dict:
        """Tạo ảnh reference cho nhân vật và địa điểm."""
        result = {"total": 0, "completed": 0, "failed": 0}

        characters = wb.get_characters()
        if not characters:
            self.log("Không có nhân vật/địa điểm trong Excel")
            return result

        # Lọc: chỉ tạo ảnh cho những cái chưa có
        pending = []
        for char in characters:
            if char.is_child:
                self.log(f"  Skip {char.id} (trẻ em)")
                continue
            if char.status and char.status.lower() in ("done", "skip"):
                self.log(f"  Skip {char.id} (status={char.status})")
                continue

            # Xác định output path
            if char.id.lower().startswith("loc"):
                img_path = self.nv_dir / f"{char.id}.png"
            else:
                img_path = self.nv_dir / f"{char.id}.png"

            # Kiểm tra file đã tồn tại và có media_id
            if img_path.exists() and char.media_id:
                self.log(f"  Skip {char.id} (đã có ảnh + media_id)")
                continue

            pending.append((char, img_path))

        result["total"] = len(pending)
        self.log(f"References cần tạo: {len(pending)}/{len(characters)} (concurrent={self.max_concurrent})")

        # Build task list
        tasks = []
        for i, (char, img_path) in enumerate(pending):
            prompt = char.english_prompt or char.vietnamese_prompt or char.name
            if not prompt:
                self.log(f"  [{i+1}/{len(pending)}] {char.id}: SKIP (không có prompt)", "WARN")
                result["failed"] += 1
                continue
            tasks.append({"idx": i, "char": char, "img_path": img_path, "prompt": prompt})

        completed_count = [0]

        def _do_char(task):
            if self._stop_flag:
                return None
            char = task["char"]
            img_path = task["img_path"]
            prompt = task["prompt"]
            idx = task["idx"]

            self.log(f"  [{idx+1}/{len(pending)}] {char.id}: {prompt[:60]}...")
            self.on_item_status("char", char.id, "running", None, {})

            def _poll_cb(info):
                self.on_item_status("char", char.id, "running", None,
                                    {"queue_pos": info.get("queue_position"),
                                     "poll_status": info.get("status")})

            t0 = time.time()
            success, media_name, server_info = self._submit_image(prompt, img_path, poll_callback=_poll_cb)
            elapsed = round(time.time() - t0, 1)

            if success:
                update_data = {"status": "done"}
                if media_name:
                    update_data["media_id"] = media_name
                wb.update_character(char.id, **update_data)
                wb.safe_save()
                completed_count[0] += 1
                self.progress("refs", completed_count[0], len(tasks), char.id)
                self.log(f"    {char.id} → OK ({elapsed}s, {server_info.get('server', '?')})")
                self.on_item_status("char", char.id, "done", str(img_path),
                                    {"elapsed": elapsed, **server_info})
                return True
            else:
                self.log(f"    {char.id} → FAIL ({elapsed}s)", "WARN")
                self.on_item_status("char", char.id, "error", None,
                                    {"elapsed": elapsed, **server_info})
                return False

        with ThreadPoolExecutor(max_workers=self.max_concurrent) as executor:
            futures = {executor.submit(_do_char, t): t for t in tasks}
            for future in as_completed(futures):
                if self._stop_flag:
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                r = future.result()
                if r is True:
                    result["completed"] += 1
                elif r is False:
                    result["failed"] += 1

        return result

    # =========================================================================
    # PHASE 2: Scene Images
    # =========================================================================

    def _generate_scenes(self, wb: PromptWorkbook) -> Dict:
        """Tạo ảnh cho TẤT CẢ scenes."""
        result = {"total": 0, "completed": 0, "failed": 0}

        scenes = wb.get_scenes()
        if not scenes:
            self.log("Không có scenes trong Excel")
            return result

        # Lọc scenes cần tạo
        pending = []
        for scene in scenes:
            if not scene.img_prompt:
                continue
            if scene.status_img and scene.status_img.lower() in ("done", "skip"):
                img_path = self.img_dir / f"scene_{scene.scene_id:03d}.png"
                if img_path.exists():
                    continue

            pending.append(scene)

        result["total"] = len(pending)
        total_scenes = len([s for s in scenes if s.img_prompt])
        self.log(f"Scenes cần tạo: {len(pending)}/{total_scenes} (concurrent={self.max_concurrent})")

        media_ids = self._load_media_ids(wb)

        completed_count = [0]

        def _do_scene(i, scene):
            if self._stop_flag:
                return None
            scene_id = scene.scene_id
            img_path = self.img_dir / f"scene_{scene_id:03d}.png"
            prompt = scene.img_prompt

            self.log(f"  [{i+1}/{len(pending)}] Scene {scene_id}: {prompt[:60]}...")
            self.on_item_status("scene", scene_id, "running", None, {})

            refs = self._build_references(scene, media_ids)
            # Log refs đang dùng
            if refs:
                ref_names = [r.name[:20] for r in refs]
                self.log(f"    Refs: {ref_names}")
            else:
                self.log(f"    Không có reference images")

            def _poll_cb(info):
                self.on_item_status("scene", scene_id, "running", None,
                                    {"queue_pos": info.get("queue_position"),
                                     "poll_status": info.get("status")})

            t0 = time.time()
            success, media_name, server_info = self._submit_image(prompt, img_path, refs, poll_callback=_poll_cb)
            elapsed = round(time.time() - t0, 1)

            if success:
                wb.update_scene(scene_id, status_img="done", img_path=str(img_path))
                if media_name:
                    wb.update_scene(scene_id, media_id=media_name)
                wb.safe_save()
                completed_count[0] += 1
                self.progress("scenes", completed_count[0], len(pending), f"scene_{scene_id:03d}")
                self.log(f"    Scene {scene_id} → OK ({elapsed}s, {server_info.get('server', '?')})")
                self.on_item_status("scene", scene_id, "done", str(img_path),
                                    {"elapsed": elapsed, **server_info})
                return True
            else:
                wb.update_scene(scene_id, status_img="error")
                wb.safe_save()
                self.log(f"    Scene {scene_id} → FAIL ({elapsed}s)", "WARN")
                self.on_item_status("scene", scene_id, "error", None,
                                    {"elapsed": elapsed, **server_info})
                return False

        with ThreadPoolExecutor(max_workers=self.max_concurrent) as executor:
            futures = {executor.submit(_do_scene, i, s): s for i, s in enumerate(pending)}
            for future in as_completed(futures):
                if self._stop_flag:
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                r = future.result()
                if r is True:
                    result["completed"] += 1
                elif r is False:
                    result["failed"] += 1

        return result

    def _load_media_ids(self, wb: PromptWorkbook) -> Dict[str, str]:
        """Load media_ids từ characters sheet."""
        media_ids = {}
        try:
            characters = wb.get_characters()
            for char in characters:
                if char.media_id:
                    media_ids[char.id] = char.media_id
                    # Cũng map theo filename
                    fname = f"{char.id}.png"
                    media_ids[fname] = char.media_id
        except Exception as e:
            self.log(f"Lỗi load media_ids: {e}", "WARN")
        self.log(f"  Media IDs loaded: {len(media_ids)}")
        return media_ids

    def _build_references(self, scene: Scene, media_ids: Dict[str, str]) -> List[ImageInput]:
        """Build ImageInput references cho scene."""
        refs = []

        # Từ reference_files (JSON list)
        ref_files = []
        if scene.reference_files:
            try:
                ref_files = json.loads(scene.reference_files) if isinstance(scene.reference_files, str) else scene.reference_files
            except (json.JSONDecodeError, TypeError):
                # Thử split by comma
                ref_files = [f.strip() for f in str(scene.reference_files).split(",") if f.strip()]

        for ref_file in ref_files:
            ref_name = ref_file.replace(".png", "").replace(".jpg", "")
            media_id = media_ids.get(ref_file) or media_ids.get(ref_name)
            if media_id:
                refs.append(ImageInput(name=media_id, input_type=ImageInputType.REFERENCE))

        # Fallback: từ characters_used + location_used nếu reference_files trống
        if not refs:
            if scene.characters_used:
                char_ids = [c.strip() for c in str(scene.characters_used).split(",") if c.strip()]
                for cid in char_ids:
                    media_id = media_ids.get(cid) or media_ids.get(f"{cid}.png")
                    if media_id:
                        refs.append(ImageInput(name=media_id, input_type=ImageInputType.REFERENCE))

            if scene.location_used:
                loc_id = scene.location_used.strip()
                media_id = media_ids.get(loc_id) or media_ids.get(f"{loc_id}.png")
                if media_id:
                    refs.append(ImageInput(name=media_id, input_type=ImageInputType.REFERENCE))

        return refs

    # =========================================================================
    # PHASE 3: Thumbnail
    # =========================================================================

    def _generate_thumbnail(self, wb: PromptWorkbook):
        """Tạo thumbnail từ nhân vật chính."""
        self.thumb_dir.mkdir(parents=True, exist_ok=True)

        characters = wb.get_characters()
        if not characters:
            self.log("Không có nhân vật để tạo thumbnail")
            return

        # Lọc bỏ locations
        actual_chars = [c for c in characters if not c.id.lower().startswith("loc") and c.role != "location"]
        if not actual_chars:
            actual_chars = characters

        # Chọn nhân vật chính (protagonist/main, không phải trẻ em)
        selected = None
        for char in actual_chars:
            if char.role and char.role.lower() in ("protagonist", "main") and not char.is_child:
                selected = char
                break

        # Fallback: nhân vật đầu tiên không phải trẻ em
        if not selected:
            for char in actual_chars:
                if not char.is_child:
                    selected = char
                    break

        # Last resort: nhân vật đầu tiên
        if not selected:
            selected = actual_chars[0]

        # Copy ảnh nhân vật vào thumb/
        src_image = self.nv_dir / f"{selected.id}.png"
        if src_image.exists():
            project_code = self.project_dir.name
            dest_image = self.thumb_dir / f"{project_code}.png"
            shutil.copy2(str(src_image), str(dest_image))
            self.log(f"Thumbnail: {selected.id} ({selected.name}) → {dest_image.name}")
        else:
            self.log(f"Ảnh {selected.id}.png chưa tồn tại, bỏ qua thumbnail", "WARN")

    # =========================================================================
    # PHASE 4: Video Generation (Image-to-Video)
    # =========================================================================

    def _generate_videos(self, wb: PromptWorkbook) -> Dict:
        """Tạo video từ ảnh scene đã có."""
        result = {"total": 0, "completed": 0, "failed": 0}

        scenes = wb.get_scenes()
        if not scenes:
            self.log("Không có scenes")
            return result

        # Lọc scene có video_prompt và ảnh đã xong
        pending = []
        for scene in scenes:
            vp = getattr(scene, 'video_prompt', '') or ''
            if not vp:
                continue
            sv = getattr(scene, 'status_vid', '') or ''
            if sv.lower() == "skip":
                continue
            if sv.lower() == "done":
                vid_path = self.vid_dir / f"scene_{scene.scene_id:03d}.mp4"
                if vid_path.exists():
                    continue
            # Cần có ảnh + media_id để làm Image-to-Video
            img_path = self.img_dir / f"scene_{scene.scene_id:03d}.png"
            media_id = getattr(scene, 'media_id', '') or ''
            if not img_path.exists() or not media_id:
                self.log(f"  Skip scene {scene.scene_id}: chưa có ảnh hoặc media_id")
                continue
            pending.append(scene)

        result["total"] = len(pending)
        if not pending:
            self.log("Không có scene nào cần tạo video")
            return result
        self.log(f"Videos cần tạo: {len(pending)} (concurrent={self.max_concurrent})")

        completed_count = [0]

        def _do_video(i, scene):
            if self._stop_flag:
                return None
            sid = scene.scene_id
            vp = scene.video_prompt
            media_id = scene.media_id
            vid_path = self.vid_dir / f"scene_{sid:03d}.mp4"

            self.log(f"  [{i+1}/{len(pending)}] Video scene {sid}: {vp[:60]}...")
            self.log(f"    media_id: {media_id[:40] if media_id else '(KHÔNG CÓ)'}")
            self.on_item_status("scene", sid, "running", None, {"phase": "video"})

            t0 = time.time()
            success, server_info = self._submit_video(vp, vid_path, media_id)
            elapsed = round(time.time() - t0, 1)

            if success:
                wb.update_scene(sid, status_vid="done", video_path=str(vid_path))
                wb.safe_save()
                completed_count[0] += 1
                self.progress("videos", completed_count[0], len(pending), f"scene_{sid:03d}")
                self.log(f"    Video scene {sid} → OK ({elapsed}s)")
                self.on_item_status("scene", sid, "done", None,
                                    {"elapsed": elapsed, "phase": "video", **server_info})
                return True
            else:
                wb.update_scene(sid, status_vid="error")
                wb.safe_save()
                self.log(f"    Video scene {sid} → FAIL ({elapsed}s)", "WARN")
                self.on_item_status("scene", sid, "error", None,
                                    {"elapsed": elapsed, "phase": "video", **server_info})
                return False

        with ThreadPoolExecutor(max_workers=self.max_concurrent) as executor:
            futures = {executor.submit(_do_video, i, s): s for i, s in enumerate(pending)}
            for future in as_completed(futures):
                if self._stop_flag:
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                r = future.result()
                if r is True:
                    result["completed"] += 1
                elif r is False:
                    result["failed"] += 1

        return result

    def _submit_video(
        self,
        prompt: str,
        output_path: Path,
        reference_image_id: str
    ) -> tuple:
        """
        Gửi prompt tạo video lên server (Image-to-Video).
        Flow 3 bước:
          1. POST /api/fix/create-video-veo3 → taskId
          2. Poll /api/fix/task-status → operations
          3. Poll Google batchCheckAsyncVideoGenerationStatus → video URL → download

        Returns:
            (success: bool, server_info: dict)
        """
        output_path.parent.mkdir(parents=True, exist_ok=True)
        sinfo = {}

        for attempt in range(self.retry_count):
            if self._stop_flag:
                return False, sinfo

            server = self.pool.pick_best_server() if self.pool else None
            if not server:
                self.log("  Chờ server...", "WARN")
                server = self.pool.wait_for_server(max_wait=300) if self.pool else None
                if not server:
                    self.log("  Không có server!", "ERROR")
                    return False, sinfo

            sinfo = {
                "server": server.name,
                "queue": server.queue_size,
                "pending": server.local_pending,
            }
            self.log(f"    → Server: {server.name} (queue={server.queue_size})")

            try:
                api = GoogleFlowAPI(
                    bearer_token=self.bearer_token,
                    project_id=self.flow_project_id,
                    timeout=self.timeout,
                    local_server_url=server.url
                )

                # Video aspect ratio
                var_str = self.config.get("flow_aspect_ratio", "landscape").upper()
                video_ar = getattr(VideoAspectRatio, var_str, VideoAspectRatio.LANDSCAPE)

                # generate_video → nội bộ sẽ:
                #   1. POST /create-video-veo3 → taskId
                #   2. _poll_proxy_video_task → poll task-status → lấy operations
                #   3. _poll_google_with_operations → poll Google → lấy video URL
                success, vid_result, error = api.generate_video(
                    prompt=prompt,
                    aspect_ratio=video_ar,
                    model=VideoModel.VEO3_I2V_FAST,
                    reference_image_id=reference_image_id
                )

                if success and vid_result and vid_result.video_url:
                    filename = output_path.stem
                    saved = api.download_video(vid_result, output_path.parent, filename)
                    if saved:
                        self.pool.mark_success(server)
                        return True, sinfo
                    else:
                        self.log(f"    [{output_path.stem}] Tải video thất bại (lần {attempt+1}/{self.retry_count})", "WARN")
                        self.pool.mark_task_failed(server)
                else:
                    err = error or "Không có video URL"
                    if "401" in err or "Authentication" in err:
                        self.log(f"    [{output_path.stem}] TOKEN HẾT HẠN — cần đổi token mới (lần {attempt+1})", "ERROR")
                    elif "400" in err or "invalid" in err.lower():
                        self.log(f"    [{output_path.stem}] GOOGLE TỪ CHỐI — {err[:300]} (lần {attempt+1})", "ERROR")
                    elif "FAILED" in err:
                        self.log(f"    [{output_path.stem}] VIDEO THẤT BẠI — {err[:300]} (lần {attempt+1})", "ERROR")
                    elif "timeout" in err.lower() or "Timeout" in err:
                        self.log(f"    [{output_path.stem}] HẾT THỜI GIAN — {err[:200]} (lần {attempt+1})", "WARN")
                    else:
                        self.log(f"    [{output_path.stem}] LỖI: {err[:300]} (lần {attempt+1})", "WARN")
                    self.pool.mark_task_failed(server)

            except Exception as e:
                self.log(f"    [{output_path.stem}] NGOẠI LỆ: {type(e).__name__}: {e} (lần {attempt+1})", "ERROR")
                if self.pool and server:
                    self.pool.mark_task_failed(server)

            if attempt < self.retry_count - 1:
                delay = 5 * (attempt + 1)
                self.log(f"    Thử lại sau {delay}s...")
                time.sleep(delay)

        return False, sinfo

    # =========================================================================
    # SERVER COMMUNICATION (Images)
    # =========================================================================

    def _submit_image(
        self,
        prompt: str,
        output_path: Path,
        refs: List[ImageInput] = None,
        poll_callback: callable = None
    ) -> tuple:
        """
        Gửi prompt tạo ảnh lên server.

        Returns:
            (success: bool, media_name: str or None, server_info: dict)
        """
        output_path.parent.mkdir(parents=True, exist_ok=True)
        sinfo = {}

        for attempt in range(self.retry_count):
            if self._stop_flag:
                return False, None, sinfo

            # Pick server
            server = self.pool.pick_best_server() if self.pool else None
            if not server:
                self.log("  Chờ server available...", "WARN")
                server = self.pool.wait_for_server(max_wait=300) if self.pool else None
                if not server:
                    self.log("  Không có server nào available!", "ERROR")
                    return False, None, sinfo

            sinfo = {
                "server": server.name,
                "server_url": server.url,
                "queue": server.queue_size,
                "pending": server.local_pending,
            }

            # Notify GUI đang gửi tới server nào
            self.log(f"    → Server: {server.name} (queue={server.queue_size}, pending={server.local_pending})")

            try:
                api = GoogleFlowAPI(
                    bearer_token=self.bearer_token,
                    project_id=self.flow_project_id,
                    timeout=self.timeout,
                    local_server_url=server.url
                )

                success, images, error = api.generate_images(
                    prompt=prompt,
                    count=1,
                    aspect_ratio=self.aspect_ratio,
                    image_inputs=refs or [],
                    poll_callback=poll_callback
                )

                if success and images:
                    img = images[0]
                    filename = output_path.stem
                    saved = api.download_image(img, output_path.parent, filename)

                    if saved:
                        self.pool.mark_success(server)
                        return True, img.media_name, sinfo
                    else:
                        self.log(f"    [{output_path.stem}] Tải ảnh thất bại (lần {attempt+1}/{self.retry_count})", "WARN")
                        self.pool.mark_task_failed(server)
                else:
                    # Phân loại lỗi rõ ràng
                    err = error or "Không rõ lỗi"
                    if "401" in err or "Authentication" in err:
                        self.log(f"    [{output_path.stem}] TOKEN HẾT HẠN — cần đổi token mới (lần {attempt+1})", "ERROR")
                    elif "400" in err or "invalid" in err.lower():
                        self.log(f"    [{output_path.stem}] GOOGLE TỪ CHỐI — {err[:200]} (lần {attempt+1})", "ERROR")
                        self.log(f"    Có thể media_id cũ không hợp lệ, thử tạo lại ảnh nhân vật trước", "WARN")
                    elif "403" in err:
                        self.log(f"    [{output_path.stem}] BỊ CHẶN (403) — {err[:200]} (lần {attempt+1})", "ERROR")
                    elif "timeout" in err.lower():
                        self.log(f"    [{output_path.stem}] HẾT THỜI GIAN CHỜ — {err[:200]} (lần {attempt+1})", "WARN")
                    else:
                        self.log(f"    [{output_path.stem}] LỖI: {err[:300]} (lần {attempt+1})", "WARN")
                    self.pool.mark_task_failed(server)

            except Exception as e:
                self.log(f"    [{output_path.stem}] NGOẠI LỆ: {type(e).__name__}: {e} (lần {attempt+1})", "ERROR")
                if self.pool and server:
                    self.pool.mark_task_failed(server)

            # Retry delay
            if attempt < self.retry_count - 1:
                delay = 2 * (attempt + 1)
                self.log(f"    Retry sau {delay}s...")
                time.sleep(delay)

        return False, None, sinfo


# =============================================================================
# CLI Entry Point
# =============================================================================

def main():
    """Chạy worker từ command line."""
    import yaml

    if len(sys.argv) < 2:
        print("Usage: python ve3_worker.py <project_dir>")
        print("  project_dir: Thư mục chứa file Excel")
        sys.exit(1)

    project_dir = sys.argv[1]

    # Load config
    config_path = VE3_DIR / "config" / "settings.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}

    worker = VE3Worker(project_dir, config)
    result = worker.run()

    if result["success"]:
        print(f"\nThành công: {result['completed']}/{result['total']} ảnh")
    else:
        print(f"\nCó lỗi: {result['completed']}/{result['total']} ảnh")
        for err in result.get("errors", []):
            print(f"  - {err}")


if __name__ == "__main__":
    main()
