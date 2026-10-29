/* lic.exe — công cụ nhỏ chạy trong bộ cài ExcelLike: lấy Mã máy và kiểm key (Ed25519, cùng định dạng với electron/license.cjs).
 *   lic.exe machine          → in Mã máy (XXXX-XXXX-XXXXX)
 *   lic.exe verify <key> | verifyfile <tệp>  → mã thoát: 0 hợp lệ, 10 sai định dạng, 11 sai chữ ký, 12 không phải ExcelLike,
 *                               13 key của máy khác, 14 hết hạn, 15 đã thu hồi; in "serial|loại|ngày cấp|hết hạn"
 * Biên dịch: x86_64-w64-mingw32-gcc -O2 -s -o lic.exe lic.c tweetnacl.c -ladvapi32 (cần pubkey.h, revoked.h)
 */
#include <windows.h>
#include <stdio.h>
#include <string.h>
#include "tweetnacl.h"
#include "pubkey.h"   /* static const unsigned char PUBKEY[32] */
#include "revoked.h"  /* static const unsigned int REVOKED[]; REVOKED_N */

void randombytes(unsigned char *x, unsigned long long n) { (void)x; (void)n; } /* không dùng (chỉ kiểm chữ ký) */

static const char ALPHA[] = "0123456789ABCDEFGHJKMNPQRSTVWXYZ";
enum { BODY = 19, SIG = 64 };

static int b32dec(const char *s, unsigned char *out, int n) {
  int bits = 0, i = 0; unsigned val = 0;
  for (; *s; s++) {
    char c = *s;
    if (c == ' ' || c == '-' || c == '_' || c == '.' || c == '\t' || c == '\r' || c == '\n') continue;
    if (c >= 'a' && c <= 'z') c -= 32;
    if (c == 'O') c = '0'; if (c == 'I' || c == 'L') c = '1';
    const char *p = strchr(ALPHA, c);
    if (!p || !c) return 0;
    val = ((val << 5) | (unsigned)(p - ALPHA)) & 0xffff; bits += 5;
    if (bits >= 8) { if (i >= n) return 0; out[i++] = (val >> (bits - 8)) & 255; bits -= 8; }
  }
  return i == n;
}
static void b32enc(const unsigned char *in, int n, char *out) {
  int bits = 0, o = 0; unsigned val = 0;
  for (int i = 0; i < n; i++) { val = (val << 8) | in[i]; bits += 8; while (bits >= 5) { out[o++] = ALPHA[(val >> (bits - 5)) & 31]; bits -= 5; } }
  if (bits > 0) out[o++] = ALPHA[(val << (5 - bits)) & 31];
  out[o] = 0;
}
/* số ngày từ 01/01/2020 của một ngày lịch */
static long days_from_civil(int y, unsigned m, unsigned d) {
  y -= m <= 2; long era = (y >= 0 ? y : y - 399) / 400; unsigned yoe = (unsigned)(y - era * 400);
  unsigned doy = (153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + d - 1; unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
  return era * 146097 + (long)doe - 719468;
}
static void ymd(long n, char *buf) {
  long z = n + days_from_civil(2020, 1, 1) + 719468; long era = (z >= 0 ? z : z - 146096) / 146097;
  unsigned doe = (unsigned)(z - era * 146097); unsigned yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
  long y = (long)yoe + era * 400; unsigned doy = doe - (365 * yoe + yoe / 4 - yoe / 100); unsigned mp = (5 * doy + 2) / 153;
  unsigned d = doy - (153 * mp + 2) / 5 + 1; unsigned m = mp + (mp < 10 ? 3 : -9);
  sprintf(buf, "%02u/%02u/%ld", d, m, y + (m <= 2));
}
static int machine_hash(unsigned char h8[8]) {
  wchar_t w[128]; DWORD sz = sizeof(w);
  if (RegGetValueW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Cryptography", L"MachineGuid", RRF_RT_REG_SZ | RRF_SUBKEY_WOW6464KEY, NULL, w, &sz) != ERROR_SUCCESS) return 0;
  char s[160] = "ExcelLike|"; int o = (int)strlen(s);
  for (int i = 0; w[i] && o < 150; i++) { wchar_t c = w[i]; if (c == ' ' || c == '\t') continue; if (c >= 'a' && c <= 'z') c -= 32; s[o++] = (char)c; }
  s[o] = 0;
  unsigned char full[64];
  crypto_hash(full, (const unsigned char *)s, (unsigned long long)o);
  memcpy(h8, full, 8);
  return 1;
}

/* tham số dòng lệnh dạng Unicode (đường dẫn có tên người dùng tiếng Việt) */
int wmain(int argc, wchar_t **wargv) {
  static char abuf[16][1024]; char *argv[16];
  if (argc > 16) argc = 16;
  for (int i = 0; i < argc; i++) { int j = 0; for (; wargv[i][j] && j < 1023; j++) abuf[i][j] = wargv[i][j] < 128 ? (char)wargv[i][j] : '?'; abuf[i][j] = 0; argv[i] = abuf[i]; }
  if (argc >= 2 && !strcmp(argv[1], "machine")) {
    unsigned char h[8]; char b[16];
    if (!machine_hash(h)) { printf("?\n"); return 2; }
    b32enc(h, 8, b);
    printf("%.4s-%.4s-%s\n", b, b + 4, b + 8);
    return 0;
  }
  if (argc >= 3 && (!strcmp(argv[1], "verify") || !strcmp(argv[1], "verifyfile"))) {
    char key[1024] = "";
    if (!strcmp(argv[1], "verifyfile")) {
      /* key đọc từ tệp (bộ cài ghi ô nhập ra tệp tạm để không vướng xuống dòng) */
      FILE *f = _wfopen(wargv[2], L"rb"); if (!f) return 10;
      size_t n = fread(key, 1, sizeof(key) - 1, f); fclose(f); key[n] = 0;
      for (size_t i = 0; i < n; i++) if ((unsigned char)key[i] < 32) key[i] = ' ';
      if (n >= 3 && (unsigned char)key[0] == 0xEF && (unsigned char)key[1] == 0xBB && (unsigned char)key[2] == 0xBF) memmove(key, key + 3, n - 2);
    } else {
      /* ghép mọi tham số còn lại (người dùng có thể dán key có khoảng trắng) */
      for (int i = 2; i < argc; i++) { if (strlen(key) + strlen(argv[i]) + 2 >= sizeof(key)) return 10; strcat(key, argv[i]); }
    }
    unsigned char all[BODY + SIG];
    if (!b32dec(key, all, BODY + SIG) || all[0] != 1 || (all[2] != 1 && all[2] != 2)) { printf("format\n"); return 10; }
    unsigned char sm[SIG + BODY], m[SIG + BODY]; unsigned long long mlen;
    memcpy(sm, all + BODY, SIG); memcpy(sm + SIG, all, BODY);
    if (crypto_sign_open(m, &mlen, sm, SIG + BODY, PUBKEY) != 0) { printf("signature\n"); return 11; }
    unsigned serial = ((unsigned)all[3] << 24) | (all[4] << 16) | (all[5] << 8) | all[6];
    long issued = (all[7] << 8) | all[8], expires = (all[9] << 8) | all[10];
    char d1[16], d2[16] = "";
    ymd(issued, d1); if (expires) ymd(expires, d2);
    printf("%u|%s|%s|%s\n", serial, all[2] == 1 ? "machine" : "shared", d1, d2);
    if (all[1] != 1) return 12;
    for (unsigned i = 0; i < REVOKED_N; i++) if (REVOKED[i] == serial) return 15;
    if (all[2] == 1) { unsigned char h[8]; if (!machine_hash(h) || memcmp(h, all + 11, 8)) return 13; }
    if (expires) { SYSTEMTIME t; GetLocalTime(&t); long today = days_from_civil(t.wYear, t.wMonth, t.wDay) - days_from_civil(2020, 1, 1); if (today > expires) return 14; }
    return 0;
  }
  printf("lic.exe machine | lic.exe verify <key> | lic.exe verifyfile <path>\n");
  return 1;
}
