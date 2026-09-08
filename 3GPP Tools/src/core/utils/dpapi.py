# --- File: src/core/utils/dpapi.py ---
"""
Core Windows DPAPI Cryptographic & Security Module.

Provides OS-level credential encryption bound to the current Windows user,
application-specific entropy salts, WinVerifyTrust Authenticode & Windows Catalog
verification, and anti-shadowing safeguards for pywin32 / crypt32.dll.
"""

import base64
import ctypes
from ctypes import wintypes
import json
import logging
import os
import re
import sys
from typing import Any, Dict, Optional, Tuple

from core.utils.paths import get_project_root

# Attempt importing Windows DPAPI securely
try:
    import win32crypt
    HAS_DPAPI = True
except ImportError:
    HAS_DPAPI = False

# ==========================================
# --- WINTRUST AUTHENTICODE & CATALOG STRUCTS ---
# ==========================================

class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", wintypes.DWORD),
        ("Data2", wintypes.WORD),
        ("Data3", wintypes.WORD),
        ("Data4", wintypes.BYTE * 8),
    ]


# Action GUID: {00AAC56B-CD44-11d0-8CC2-00C04FC295EE} (WINTRUST_ACTION_GENERIC_VERIFY_V2)
WINTRUST_ACTION_GENERIC_VERIFY_V2 = GUID(
    0x00AAC56B,
    0xCD44,
    0x11D0,
    (wintypes.BYTE * 8)(0x8C, 0xC2, 0x00, 0xC0, 0x4F, 0xC2, 0x95, 0xEE)
)

WTD_CHOICE_FILE = 1
WTD_UI_NONE = 2
WTD_REVOKE_NONE = 0
WTD_STATEACTION_IGNORE = 0
TRUST_E_NOSIGNATURE = 0x800B0100


class WINTRUST_FILE_INFO(ctypes.Structure):
    _fields_ = [
        ("cbStruct", wintypes.DWORD),
        ("pcwszFilePath", wintypes.LPCWSTR),
        ("hFile", wintypes.HANDLE),
        ("pgKnownSubject", ctypes.c_void_p),
    ]


class WINTRUST_DATA(ctypes.Structure):
    _fields_ = [
        ("cbStruct", wintypes.DWORD),
        ("pPolicyCallbackData", ctypes.c_void_p),
        ("pSIPClientData", ctypes.c_void_p),
        ("dwUIChoice", wintypes.DWORD),
        ("fdwRevocationChecks", wintypes.DWORD),
        ("dwUnionChoice", wintypes.DWORD),
        ("pFile", ctypes.POINTER(WINTRUST_FILE_INFO)),
        ("dwStateAction", wintypes.DWORD),
        ("hWVTStateData", wintypes.HANDLE),
        ("pwszURLReference", wintypes.LPCWSTR),
        ("dwProvFlags", wintypes.DWORD),
        ("dwUIContext", wintypes.DWORD),
        ("pSignatureSettings", ctypes.c_void_p),
    ]


# ==========================================
# --- SECURITY EXCEPTIONS & HELPERS ---
# ==========================================

class SecurityIntegrityError(Exception):
    """Raised when a core cryptographic binary fails provenance or signature checks."""
    pass


def redact_sensitive_urls(message: str) -> str:
    """Removes passwords and secrets from embedded URLs in error and exception strings."""
    if not message:
        return ""
    return re.sub(r"(://[^:/@\s]+):([^@\s]+)@", r"\1:***@", str(message))


# ==========================================
# --- SECURE CREDENTIAL STORE (DPAPI) ---
# ==========================================

class SecureCredentialStore:
    """
    Centralized credential manager leveraging Windows DPAPI with runtime
    Authenticode, Windows Catalog, and 64-bit module provenance validation.
    """

    _verified = False
    _DEFAULT_ENTROPY = b"3GPP_Tools_Telekom_Entropy_V1"

    @staticmethod
    def is_available() -> bool:
        """Returns True if the underlying Windows DPAPI subsystem is available."""
        return HAS_DPAPI

    @classmethod
    def _get_system_directory(cls) -> str:
        """Dynamically queries the Windows kernel for the genuine system folder path."""
        try:
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.GetSystemDirectoryW.argtypes = [wintypes.LPWSTR, wintypes.UINT]
            kernel32.GetSystemDirectoryW.restype = wintypes.UINT

            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH + 1)
            length = kernel32.GetSystemDirectoryW(buf, len(buf))
            if length > 0:
                return os.path.normpath(buf.value)
        except Exception:
            pass
        return os.path.normpath(os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32"))

    @classmethod
    def verify_library_integrity(cls) -> None:
        """
        Validates that win32crypt and crypt32.dll are authentic Microsoft system binaries
        and have not been shadowed, intercepted, or hijacked from local directories.
        """
        if cls._verified:
            return

        if not HAS_DPAPI:
            raise SecurityIntegrityError("Windows DPAPI (pywin32) is not installed.")

        # 1. Check Python Module Provenance (Prevents local directory shadowing)
        module_file = getattr(win32crypt, "__file__", "")
        if not module_file:
            raise SecurityIntegrityError("win32crypt has no physical module path.")

        norm_path = os.path.normpath(os.path.abspath(module_file)).lower()
        project_root = os.path.normpath(str(get_project_root())).lower()

        if norm_path.startswith(project_root):
            raise SecurityIntegrityError(
                f"Security violation: win32crypt was loaded from the local workspace ({norm_path}) "
                f"instead of Python site-packages."
            )

        # 2. Verify that the loaded crypt32.dll module is executing out of the genuine system directory
        system_dir = cls._get_system_directory()
        system_crypt32 = os.path.join(system_dir, "crypt32.dll")

        if not cls._verify_loaded_module(system_crypt32):
            raise SecurityIntegrityError(f"crypt32.dll in memory was not loaded from {system_dir}.")

        # 3. Verify on-disk file existence in the system directory
        if not os.path.exists(system_crypt32):
            raise SecurityIntegrityError(f"Critical system binary crypt32.dll not found in {system_dir}.")

        # 4. Verify Digital Signature (Authenticode + Catalog Fallback)
        if not cls._verify_authenticode(system_crypt32):
            raise SecurityIntegrityError(
                f"Digital signature verification failed for {system_crypt32}. "
                "The binary signature is untrusted, invalid, or modified."
            )

        cls._verified = True
        logging.info("🛡️ Cryptographic subsystem integrity verified: crypt32.dll is authentic.")

    @classmethod
    def _verify_loaded_module(cls, expected_crypt32_path: str) -> bool:
        """
        Verifies that crypt32.dll in process memory is physically mapped
        from the expected Windows system directory using explicit 64-bit typing.
        """
        try:
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

            # Explicitly type 64-bit return values and parameters
            kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
            kernel32.GetModuleHandleW.restype = wintypes.HMODULE

            kernel32.LoadLibraryExW.argtypes = [wintypes.LPCWSTR, wintypes.HANDLE, wintypes.DWORD]
            kernel32.LoadLibraryExW.restype = wintypes.HMODULE

            kernel32.GetModuleFileNameW.argtypes = [wintypes.HMODULE, wintypes.LPWSTR, wintypes.DWORD]
            kernel32.GetModuleFileNameW.restype = wintypes.DWORD

            hModule = kernel32.GetModuleHandleW("crypt32.dll")
            if not hModule:
                # Pin directly to system32 if not already in memory
                hModule = kernel32.LoadLibraryExW(expected_crypt32_path, None, 0x00000800)  # LOAD_LIBRARY_SEARCH_SYSTEM32

            if not hModule:
                return False

            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH + 1)
            length = kernel32.GetModuleFileNameW(hModule, buf, len(buf))
            if length == 0:
                return False

            loaded_path = os.path.normpath(buf.value)
            expected_path = os.path.normpath(expected_crypt32_path)

            # Normalize device path prefixes if present
            if loaded_path.startswith("\\\\?\\"):
                loaded_path = loaded_path[4:]
            if expected_path.startswith("\\\\?\\"):
                expected_path = expected_path[4:]

            # Check if both paths point to the exact same physical inode on disk
            if os.path.exists(loaded_path) and os.path.exists(expected_path):
                try:
                    if os.path.samefile(loaded_path, expected_path):
                        return True
                except Exception:
                    pass

            return loaded_path.lower() == expected_path.lower()
        except Exception as e:
            logging.error(f"Failed to verify loaded module: {e}")
            return False

    @classmethod
    def _verify_authenticode(cls, file_path: str) -> bool:
        """
        Verifies that the binary has a valid digital signature.
        Supports both embedded Authenticode signatures and Windows Security Catalog signatures.
        """
        try:
            wintrust = ctypes.WinDLL("wintrust", use_last_error=True)
            WinVerifyTrust = wintrust.WinVerifyTrust
            WinVerifyTrust.argtypes = [wintypes.HWND, ctypes.c_void_p, ctypes.c_void_p]
            WinVerifyTrust.restype = wintypes.LONG

            file_info = WINTRUST_FILE_INFO(
                cbStruct=ctypes.sizeof(WINTRUST_FILE_INFO),
                pcwszFilePath=file_path,
                hFile=None,
                pgKnownSubject=None,
            )

            wintrust_data = WINTRUST_DATA(
                cbStruct=ctypes.sizeof(WINTRUST_DATA),
                pPolicyCallbackData=None,
                pSIPClientData=None,
                dwUIChoice=WTD_UI_NONE,
                fdwRevocationChecks=WTD_REVOKE_NONE,
                dwUnionChoice=WTD_CHOICE_FILE,
                pFile=ctypes.pointer(file_info),
                dwStateAction=WTD_STATEACTION_IGNORE,
                hWVTStateData=None,
                pwszURLReference=None,
                dwProvFlags=0x00000080,  # WTD_CACHE_ONLY_URL_RETRIEVAL
                dwUIContext=0,
                pSignatureSettings=None,
            )

            status = WinVerifyTrust(
                None,
                ctypes.byref(WINTRUST_ACTION_GENERIC_VERIFY_V2),
                ctypes.byref(wintrust_data),
            )

            # Valid embedded signature
            if status == 0:
                return True

            # If no embedded signature is present (standard for core OS binaries), verify Catalog
            unsigned_status = status & 0xFFFFFFFF
            if unsigned_status == TRUST_E_NOSIGNATURE:
                return cls._verify_catalog_signature(file_path)

            return False
        except Exception as e:
            logging.error(f"Authenticode check error: {e}")
            return False

    @staticmethod
    def _verify_catalog_signature(file_path: str) -> bool:
        """Verifies that a file's hash is registered in the official Windows Security Catalog (CatRoot)."""
        try:
            wintrust = ctypes.WinDLL("wintrust", use_last_error=True)
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

            # Configure 64-bit function types to prevent handle truncation
            wintrust.CryptCATAdminAcquireContext.argtypes = [
                ctypes.POINTER(wintypes.HANDLE),
                ctypes.c_void_p,
                wintypes.DWORD,
            ]
            wintrust.CryptCATAdminAcquireContext.restype = wintypes.BOOL

            wintrust.CryptCATAdminCalcHashFromFileHandle.argtypes = [
                wintypes.HANDLE,
                ctypes.POINTER(wintypes.DWORD),
                ctypes.c_void_p,
                wintypes.DWORD,
            ]
            wintrust.CryptCATAdminCalcHashFromFileHandle.restype = wintypes.BOOL

            wintrust.CryptCATAdminEnumCatalogFromHash.argtypes = [
                wintypes.HANDLE,
                ctypes.c_void_p,
                wintypes.DWORD,
                wintypes.DWORD,
                ctypes.POINTER(wintypes.HANDLE),
            ]
            wintrust.CryptCATAdminEnumCatalogFromHash.restype = wintypes.HANDLE

            wintrust.CryptCATAdminReleaseCatalogContext.argtypes = [
                wintypes.HANDLE,
                wintypes.HANDLE,
                wintypes.DWORD,
            ]
            wintrust.CryptCATAdminReleaseCatalogContext.restype = wintypes.BOOL

            wintrust.CryptCATAdminReleaseContext.argtypes = [
                wintypes.HANDLE,
                wintypes.DWORD,
            ]
            wintrust.CryptCATAdminReleaseContext.restype = wintypes.BOOL

            kernel32.CreateFileW.argtypes = [
                wintypes.LPCWSTR,
                wintypes.DWORD,
                wintypes.DWORD,
                ctypes.c_void_p,
                wintypes.DWORD,
                wintypes.DWORD,
                wintypes.HANDLE,
            ]
            kernel32.CreateFileW.restype = wintypes.HANDLE

            kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
            kernel32.CloseHandle.restype = wintypes.BOOL

            hCatAdmin = wintypes.HANDLE()
            if not wintrust.CryptCATAdminAcquireContext(ctypes.byref(hCatAdmin), None, 0):
                return False

            try:
                GENERIC_READ = 0x80000000
                FILE_SHARE_READ = 0x00000001
                FILE_SHARE_WRITE = 0x00000002
                OPEN_EXISTING = 3
                FILE_FLAG_SEQUENTIAL_SCAN = 0x08000000

                hFile = kernel32.CreateFileW(
                    file_path,
                    GENERIC_READ,
                    FILE_SHARE_READ | FILE_SHARE_WRITE,
                    None,
                    OPEN_EXISTING,
                    FILE_FLAG_SEQUENTIAL_SCAN,
                    None,
                )
                if hFile == -1 or hFile == wintypes.HANDLE(-1).value:
                    return False

                try:
                    cbHash = wintypes.DWORD(0)
                    wintrust.CryptCATAdminCalcHashFromFileHandle(
                        hFile, ctypes.byref(cbHash), None, 0
                    )
                    if cbHash.value == 0:
                        return False

                    pbHash = (ctypes.c_byte * cbHash.value)()
                    if not wintrust.CryptCATAdminCalcHashFromFileHandle(
                        hFile, ctypes.byref(cbHash), pbHash, 0
                    ):
                        return False

                    hCatInfo = wintrust.CryptCATAdminEnumCatalogFromHash(
                        hCatAdmin, pbHash, cbHash.value, 0, None
                    )
                    if hCatInfo:
                        wintrust.CryptCATAdminReleaseCatalogContext(hCatAdmin, hCatInfo, 0)
                        return True
                    return False
                finally:
                    kernel32.CloseHandle(hFile)
            finally:
                wintrust.CryptCATAdminReleaseContext(hCatAdmin, 0)
        except Exception as e:
            logging.error(f"Catalog signature verification error: {e}")
            return False

    @classmethod
    def encrypt(
        cls,
        plain_text: str,
        entropy: Optional[bytes] = None,
        description: str = "3GPP_Tools_Encrypted_Secret"
    ) -> Tuple[bool, str]:
        """
        Encrypts a plaintext string using Windows DPAPI bound to the current user.
        Returns: (success: bool, encrypted_base64: str)
        """
        if not plain_text:
            return True, ""

        try:
            cls.verify_library_integrity()
            salt = entropy if entropy is not None else cls._DEFAULT_ENTROPY
            encrypted_bytes = win32crypt.CryptProtectData(
                plain_text.encode("utf-8"),
                description,
                salt,
                None,
                None,
                0
            )
            return True, base64.b64encode(encrypted_bytes).decode("ascii")
        except SecurityIntegrityError as e:
            logging.critical(f"Aborting encryption due to integrity failure: {e}")
            return False, ""
        except Exception as e:
            logging.error(f"Failed to encrypt credentials via DPAPI: {redact_sensitive_urls(str(e))}")
            return False, ""

    @classmethod
    def decrypt(
        cls,
        cipher_text: str,
        entropy: Optional[bytes] = None
    ) -> str:
        """
        Decrypts a Base64 DPAPI ciphertext with automatic fallback for older profiles.
        Returns plaintext string on success, or empty string if decryption fails.
        """
        if not cipher_text:
            return ""

        try:
            cls.verify_library_integrity()
            raw_bytes = base64.b64decode(cipher_text.encode("ascii"))
            salt = entropy if entropy is not None else cls._DEFAULT_ENTROPY

            # Attempt 1: Decrypt using application-level entropy salt
            try:
                _, decrypted_bytes = win32crypt.CryptUnprotectData(
                    raw_bytes,
                    salt,
                    None,
                    None,
                    0
                )
                return decrypted_bytes.decode("utf-8")
            except Exception:
                # Attempt 2: Backward-compatibility fallback for profiles created without salt
                if salt is not None:
                    _, decrypted_bytes = win32crypt.CryptUnprotectData(
                        raw_bytes,
                        None,
                        None,
                        None,
                        0
                    )
                    return decrypted_bytes.decode("utf-8")
                raise
        except SecurityIntegrityError as e:
            logging.critical(f"Aborting decryption due to integrity failure: {e}")
            return ""
        except Exception as e:
            logging.warning(f"Could not decrypt stored secret: {redact_sensitive_urls(str(e))}")
            return ""

    @classmethod
    def encrypt_dict(
        cls,
        data: Dict[str, Any],
        entropy: Optional[bytes] = None,
        description: str = "3GPP_Tools_Encrypted_Payload"
    ) -> Tuple[bool, str]:
        """Serializes a dictionary to JSON and encrypts it via DPAPI."""
        try:
            serialized = json.dumps(data)
            return cls.encrypt(serialized, entropy=entropy, description=description)
        except Exception as e:
            logging.error(f"Failed to serialize and encrypt payload: {redact_sensitive_urls(str(e))}")
            return False, ""

    @classmethod
    def decrypt_dict(
        cls,
        cipher_text: str,
        entropy: Optional[bytes] = None
    ) -> Dict[str, Any]:
        """Decrypts a DPAPI ciphertext and parses it as a JSON dictionary."""
        plain = cls.decrypt(cipher_text, entropy=entropy)
        if not plain:
            return {}
        try:
            parsed = json.loads(plain)
            return parsed if isinstance(parsed, dict) else {}
        except Exception as e:
            logging.warning(f"Failed to parse decrypted dictionary payload: {redact_sensitive_urls(str(e))}")
            return {}