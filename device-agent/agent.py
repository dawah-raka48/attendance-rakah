#!/usr/bin/env python3
"""
Attendance Rakah - ZKTeco MB1000/ID sync agent.
Reads users and raw attendance events only. It never enrolls users,
changes device settings, or deletes attendance records.
"""
import getpass
import json
import os
import time
import urllib.request
import urllib.error
from pathlib import Path
from zk import ZK

DEVICE_IP = "192.168.100.76"
DEVICE_PORT = 4370
DEVICE_PASSWORD = 0
API_URL = "https://script.google.com/macros/s/AKfycbzpVs-c8cVC77s6H3KQNMoxgz9n5_z3W5hg3i47RT8kFMzoek_wkJlR9Xn0TZSuiQeUPQ/exec"
POLL_SECONDS = 60
BATCH_SIZE = 250
CONFIG_PATH = Path(__file__).with_name("config.json")


def load_key():
    if CONFIG_PATH.exists():
        cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        key = str(cfg.get("device_sync_key", "")).strip()
        if key:
            return key
    print("أدخل مفتاح المزامنة الذي أعددته في Apps Script (لن يظهر أثناء الكتابة):")
    key = getpass.getpass("Device sync key: ").strip()
    if len(key) < 32:
        raise RuntimeError("مفتاح المزامنة غير صالح أو قصير.")
    CONFIG_PATH.write_text(
        json.dumps({"device_sync_key": key}, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    try:
        os.chmod(CONFIG_PATH, 0o600)
    except OSError:
        pass
    print("تم حفظ المفتاح محليًا في config.json. لا ترفع هذا الملف إلى GitHub.")
    return key


def post_batch(key, users, records):
    payload = {
        "action": "syncDeviceData",
        "deviceKey": key,
        "users": users,
        "records": records,
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        API_URL,
        data=body,
        headers={"Content-Type": "text/plain;charset=utf-8"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=45) as response:
        result = json.loads(response.read().decode("utf-8"))
    if not result.get("ok"):
        raise RuntimeError(result.get("error", "API returned an unknown error"))
    return result.get("data", {})


def read_device():
    zk = ZK(
        DEVICE_IP,
        port=DEVICE_PORT,
        timeout=20,
        password=DEVICE_PASSWORD,
        force_udp=False,
    )
    conn = None
    try:
        conn = zk.connect()
        device_users = conn.get_users()
        attendance = conn.get_attendance()
        users = [
            {
                "userId": str(u.user_id),
                "deviceUid": str(u.uid),
                "deviceName": str(u.name or ""),
            }
            for u in device_users
        ]
        records = [
            {
                "userId": str(r.user_id),
                "deviceUid": str(getattr(r, "uid", "")),
                "timestamp": r.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                "status": str(getattr(r, "status", "")),
                "punch": str(getattr(r, "punch", "")),
            }
            for r in attendance
        ]
        return users, records
    finally:
        if conn:
            conn.disconnect()


def sync_once(key):
    users, records = read_device()
    totals = {"employeesAdded": 0, "recordsInserted": 0, "duplicates": 0}
    batches = [records[i:i + BATCH_SIZE] for i in range(0, len(records), BATCH_SIZE)]
    if not batches:
        batches = [[]]
    for index, batch in enumerate(batches):
        result = post_batch(key, users if index == 0 else [], batch)
        for field in totals:
            totals[field] += int(result.get(field, 0))
    print(
        f"المستخدمون على الجهاز: {len(users)} | "
        f"السجلات المقروءة: {len(records)} | "
        f"موظفون أُضيفوا: {totals['employeesAdded']} | "
        f"سجلات جديدة: {totals['recordsInserted']} | "
        f"مكررة تم تجاهلها: {totals['duplicates']}"
    )


def main():
    key = load_key()
    print(f"بدء مزامنة جهاز ZKTeco {DEVICE_IP}:{DEVICE_PORT} كل {POLL_SECONDS} ثانية.")
    print("وضع القراءة فقط: لا تعديل ولا حذف لسجلات الجهاز. أوقف البرنامج بـ Ctrl+C.")
    while True:
        try:
            sync_once(key)
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")[:500]
            print(f"خطأ HTTP: {exc.code} - {detail}")
        except Exception as exc:
            print(f"تعذرت المزامنة: {exc}")
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nتم إيقاف برنامج المزامنة.")
