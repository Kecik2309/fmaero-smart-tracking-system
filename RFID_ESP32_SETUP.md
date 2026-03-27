# ESP32 RC522 RFID Setup

Fail ini sambungkan `ESP32 + RC522 + LCD I2C 16x2` terus ke Firebase supaya web app ini auto-terima scan RFID.
Ia juga menyokong 2 butang fizikal untuk pilih `Stock IN` atau `Stock OUT` sebelum scan tag.

## Fail yang digunakan

- Sketch: [esp32_rc522_firebase.ino](/c:/FYP_FMAERO_PROJECT/project%20fyp/esp32_rc522_firebase.ino)
- Web app listener: [app.js](/c:/FYP_FMAERO_PROJECT/project%20fyp/app.js#L3746)

## Wiring RC522 ke ESP32

- `SDA / SS` -> `GPIO 5`
- `SCK` -> `GPIO 18`
- `MOSI` -> `GPIO 23`
- `MISO` -> `GPIO 19`
- `RST` -> `GPIO 4`
- `3.3V` -> `3V3`
- `GND` -> `GND`

## Wiring LCD I2C 16x2 ke ESP32

- `VCC` -> `5V` atau `VIN` ikut modul LCD anda
- `GND` -> `GND`
- `SDA` -> `GPIO 21`
- `SCL` -> `GPIO 22`

## Wiring 2 Button ke ESP32

- Butang `Stock IN`:
  - satu kaki -> `GPIO 25`
  - satu kaki -> `GND`
- Butang `Stock OUT`:
  - satu kaki -> `GPIO 26`
  - satu kaki -> `GND`

Sketch guna `INPUT_PULLUP`, jadi anda tak perlu resistor luaran untuk setup asas ini.

Kalau board anda guna pin lain, tukar constant pin dalam sketch.

## Library Arduino yang perlu install

- `MFRC522` by GithubCommunity / Miguel Balboa
- `LiquidCrystal_I2C`

ESP32 core biasanya sudah ada:

- `WiFi.h`
- `HTTPClient.h`
- `WiFiClientSecure.h`
- `SPI.h`

## Konfigurasi dalam sketch

Tukar nilai ini:

```cpp
const char* WIFI_SSID = "YOUR_WIFI_NAME";
const char* WIFI_PASSWORD = "YOUR_WIFI_PASSWORD";
const char* FIREBASE_AUTH = "";
```

`FIREBASE_AUTH`:

- Biarkan kosong jika Firebase RTDB anda benarkan write tanpa token untuk device ini.
- Isi token / database secret jika rules Firebase anda perlukan auth.

## Cara sistem ini bekerja

ESP32 akan tulis scan ke:

```text
rfidScans/latest
```

Payload yang dihantar:

```json
{
  "tag": "04A1B2C3D4",
  "action": "IN",
  "source": "esp32-rc522",
  "scannedAt": "boot-ms-123456"
}
```

Web app anda sekarang dengar path ini dan akan terus proses tag tersebut, termasuk pilih `IN` atau `OUT` ikut butang yang ditekan.

## Apa yang LCD akan paparkan

- Status boot
- Status sambungan Wi-Fi
- Mod semasa `IN` atau `OUT`
- UID tag yang dikesan
- Status upload ke Firebase
- Ralat ringkas jika upload gagal

## Cara test

1. Upload sketch ke ESP32.
2. Buka Serial Monitor pada `115200`.
3. Pastikan web app dibuka.
4. Dalam sistem web, isi `RFID Tag` pada material dengan UID yang sama seperti keluar di Serial Monitor.
5. Tekan butang `Stock IN` atau `Stock OUT`.
6. Tap tag pada RC522.
7. Web app akan buka flow transaksi untuk material itu dengan type yang sudah dipilih ikut butang.

## Kalau scan tak masuk

- Semak Serial Monitor, tengok sama ada tag berjaya dibaca.
- Semak Wi-Fi ESP32 betul-betul connected.
- Semak Firebase RTDB rules benarkan write ke `rfidScans/latest`.
- Semak nilai `RFID Tag` dalam material sama tepat dengan UID tag.
- Semak web app dibuka pada page inventory utama.

## Nota penting

- RC522 guna `3.3V`, bukan `5V`.
- LCD I2C biasa selalunya guna alamat `0x27`. Jika skrin anda tak menyala, semak alamat I2C dan ubah `LCD_I2C_ADDRESS` dalam sketch.
- UID tag yang dibaca sketch ditukar ke huruf besar. Dalam web app, tag juga dinormalkan ke huruf besar, jadi perbandingan kekal konsisten.
- Sketch ini guna HTTPS dan `setInsecure()` untuk mudahkan test. Untuk production, lebih baik guna certificate validation yang betul.
