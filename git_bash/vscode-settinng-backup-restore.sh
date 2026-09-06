#!/usr/bin/env bash

# ============================================================
# VS Code settings.json 백업 / 복원 / 비교
# ============================================================

SETTINGS="$HOME/AppData/Roaming/Code/User/settings.json"
BACKUP="./vscode-settings-backup.json"
BEFORE_RESTORE="./vscode-settings-before-restore.json"

# ------------------------------------------------------------
# 사용법
# ------------------------------------------------------------

usage() {
    SCRIPT_NAME="./$(basename "$0")"
    echo
    echo "사용법:"
    echo
    echo "  $SCRIPT_NAME backup          settings.json 백업"
    echo "  $SCRIPT_NAME diff            현재 설정과 백업 비교"
    echo "  $SCRIPT_NAME restore         백업에서 복원"
    echo
    read -r -p "계속하려면 Enter 키를 누르세요..."
}

# ------------------------------------------------------------
# settings.json 존재 확인
# ------------------------------------------------------------

check_settings() {
    if [ ! -f "$SETTINGS" ]; then
        echo "오류: settings.json을 찾을 수 없습니다."
        echo
        echo "확인한 경로:"
        echo "$SETTINGS"
        exit 1
    fi
}

# ------------------------------------------------------------
# 백업
# ------------------------------------------------------------

backup() {

    check_settings

    cp "$SETTINGS" "$BACKUP"

    echo
    echo "백업 완료"
    echo
    echo "원본 : $SETTINGS"
    echo "백업 : $BACKUP"
}

# ------------------------------------------------------------
# 비교
# ------------------------------------------------------------

diff_settings() {

    check_settings

    if [ ! -f "$BACKUP" ]; then
        echo "오류: 백업 파일이 없습니다."
        echo "$BACKUP"
        exit 1
    fi

    echo
    echo "현재 설정:"
    echo "  $SETTINGS"
    echo
    echo "백업 설정:"
    echo "  $BACKUP"
    echo
    echo "============================================================"
    echo

    diff -u "$BACKUP" "$SETTINGS" || true
}

# ------------------------------------------------------------
# 복원
# ------------------------------------------------------------

restore() {

    check_settings

    if [ ! -f "$BACKUP" ]; then
        echo "오류: 백업 파일이 없습니다."
        echo "$BACKUP"
        exit 1
    fi

    echo
    echo "백업 파일:"
    echo "  $BACKUP"
    echo
    echo "현재 설정을 먼저 다음 파일로 보관합니다:"
    echo "  $BEFORE_RESTORE"
    echo

    read -r -p "복원하시겠습니까? [y/N] " ANSWER

    if [[ ! "$ANSWER" =~ ^[Yy]$ ]]; then
        echo "복원을 취소했습니다."
        exit 0
    fi

    # 복원 전 현재 설정 백업
    cp "$SETTINGS" "$BEFORE_RESTORE"

    # 백업에서 복원
    cp "$BACKUP" "$SETTINGS"

    echo
    echo "복원 완료"
    echo
    echo "복원된 설정:"
    echo "  $SETTINGS"
    echo
    echo "복원 전 설정:"
    echo "  $BEFORE_RESTORE"
}

# ------------------------------------------------------------
# 메인
# ------------------------------------------------------------

case "${1:-}" in

    backup)
        backup
        ;;

    diff)
        diff_settings
        ;;

    restore)
        restore
        ;;

    *)
        usage
        ;;

esac