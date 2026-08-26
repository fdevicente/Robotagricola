import os, pytest


@pytest.fixture
def setup_dirs(tmp_path):
    master = tmp_path / "master.xlsx"
    master.write_text("fake excel")
    backup_dir = tmp_path / "Backups"
    return str(master), str(backup_dir)


def test_backup_master_creates_current(setup_dirs):
    master, backup_dir = setup_dirs
    from infrastructure.backups import backup_master
    backup_master(reason="test", excel_path=master, backup_base=backup_dir,
                  cola_path=os.path.join(backup_dir, "cola_test.jsonl"))
    current = os.path.join(backup_dir, "Master", "current.xlsx")
    assert os.path.exists(current)


def test_backup_master_creates_snapshot(setup_dirs):
    master, backup_dir = setup_dirs
    from infrastructure.backups import backup_master
    backup_master(reason="test", excel_path=master, backup_base=backup_dir,
                  cola_path=os.path.join(backup_dir, "cola_test.jsonl"))
    snapshots = os.path.join(backup_dir, "Master", "snapshots")
    files = os.listdir(snapshots)
    assert len(files) == 1
    assert files[0].endswith(".xlsx")


def test_daily_snapshot_rotates_30_days(setup_dirs):
    master, backup_dir = setup_dirs
    from infrastructure.backups import backup_master
    snapshots_dir = os.path.join(backup_dir, "Master", "snapshots")
    os.makedirs(snapshots_dir, exist_ok=True)
    for i in range(35):
        f = os.path.join(snapshots_dir, f"2026-04-{i:02d}_18-00.xlsx")
        open(f, 'w').close()
    backup_master(reason="rotation", excel_path=master, backup_base=backup_dir,
                  cola_path=os.path.join(backup_dir, "cola_test.jsonl"))
    files = os.listdir(snapshots_dir)
    assert len(files) <= 31


def test_backup_codebase(setup_dirs, tmp_path):
    _, backup_dir = setup_dirs
    robot_dir = tmp_path / "Robot"
    robot_dir.mkdir()
    (robot_dir / "main.py").write_text("print('hi')")
    (robot_dir / "config.py").write_text("X=1")
    from infrastructure.backups import backup_codebase
    backup_codebase(robot_dir=str(robot_dir), backup_base=backup_dir)
    dest = os.path.join(backup_dir, "Robot", "current", "main.py")
    assert os.path.exists(dest)
