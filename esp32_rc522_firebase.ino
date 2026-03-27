#include <WiFi.h>
#include <WiFiClientSecure.h>
#include <HTTPClient.h>
#include <SPI.h>
#include <MFRC522.h>
#include <Wire.h>
#include <LiquidCrystal_I2C.h>

// Wi-Fi settings
const char* WIFI_SSID = "TOKEY MINA";
const char* WIFI_PASSWORD = "minatergojes";

// Firebase RTDB endpoint used by the web app
const char* FIREBASE_RTDB_HOST = "https://fmaero-smart-tracking-system-default-rtdb.asia-southeast1.firebasedatabase.app";

// Optional:
// If your Firebase Realtime Database requires auth, put your database secret or token here.
// Leave empty if your rules currently allow writes from this device.
const char* FIREBASE_AUTH = "";

// RFID scan write target expected by the web app
const char* FIREBASE_SCAN_PATH = "/rfidScans/latest.json";

// ESP32 + RC522 pin mapping
// Change these pins if your hardware uses a different wiring layout.
constexpr uint8_t RFID_SS_PIN = 5;   // SDA / SS
constexpr uint8_t RFID_RST_PIN = 4;  // RST
constexpr uint8_t RFID_SCK_PIN = 18;
constexpr uint8_t RFID_MISO_PIN = 19;
constexpr uint8_t RFID_MOSI_PIN = 23;

// I2C LCD settings
constexpr uint8_t LCD_I2C_ADDRESS = 0x27;
constexpr uint8_t LCD_COLUMNS = 16;
constexpr uint8_t LCD_ROWS = 2;
constexpr uint8_t LCD_SDA_PIN = 21;
constexpr uint8_t LCD_SCL_PIN = 22;

// Physical buttons for stock flow selection
constexpr uint8_t BUTTON_STOCK_IN_PIN = 25;
constexpr uint8_t BUTTON_STOCK_OUT_PIN = 26;

MFRC522 mfrc522(RFID_SS_PIN, RFID_RST_PIN);
WiFiClientSecure secureClient;
LiquidCrystal_I2C lcd(LCD_I2C_ADDRESS, LCD_COLUMNS, LCD_ROWS);

String lastTag = "";
String selectedAction = "IN";
unsigned long lastSentAt = 0;
const unsigned long DUPLICATE_BLOCK_MS = 1500;
const unsigned long WIFI_CONNECT_TIMEOUT_MS = 15000;
unsigned long lastButtonPressAt = 0;
const unsigned long BUTTON_DEBOUNCE_MS = 220;

void showLcdMessage(const String& line1, const String& line2 = "") {
  lcd.clear();
  lcd.setCursor(0, 0);
  lcd.print(line1.substring(0, LCD_COLUMNS));
  lcd.setCursor(0, 1);
  lcd.print(line2.substring(0, LCD_COLUMNS));
}

String buildScanUrl() {
  String url = String(FIREBASE_RTDB_HOST) + FIREBASE_SCAN_PATH;
  if (String(FIREBASE_AUTH).length() > 0) {
    url += "?auth=" + String(FIREBASE_AUTH);
  }
  return url;
}

String getIsoTimestamp() {
  // If you want a real UTC timestamp, add NTP sync.
  // This fallback still gives the web app a unique scan marker.
  unsigned long ms = millis();
  return String("boot-ms-") + String(ms);
}

String readTagUid() {
  String uid = "";
  for (byte i = 0; i < mfrc522.uid.size; i++) {
    if (mfrc522.uid.uidByte[i] < 0x10) {
      uid += "0";
    }
    uid += String(mfrc522.uid.uidByte[i], HEX);
  }
  uid.toUpperCase();
  return uid;
}

void updateSelectedAction(const String& nextAction) {
  selectedAction = nextAction == "OUT" ? "OUT" : "IN";
  Serial.print("Selected action: ");
  Serial.println(selectedAction);
  showLcdMessage("Mode: " + selectedAction, "Tap RFID tag");
}

void handleButtons() {
  const unsigned long now = millis();
  if (now - lastButtonPressAt < BUTTON_DEBOUNCE_MS) {
    return;
  }

  if (digitalRead(BUTTON_STOCK_IN_PIN) == LOW && selectedAction != "IN") {
    lastButtonPressAt = now;
    updateSelectedAction("IN");
    return;
  }

  if (digitalRead(BUTTON_STOCK_OUT_PIN) == LOW && selectedAction != "OUT") {
    lastButtonPressAt = now;
    updateSelectedAction("OUT");
  }
}

bool sendTagToFirebase(const String& tag, const String& action) {
  if (WiFi.status() != WL_CONNECTED) {
    Serial.println("WiFi disconnected. Reconnecting...");
    showLcdMessage("WiFi reconnect", "Please wait...");
    WiFi.begin(WIFI_SSID, WIFI_PASSWORD);
    unsigned long startedAt = millis();
    while (WiFi.status() != WL_CONNECTED && millis() - startedAt < 10000) {
      delay(300);
      Serial.print(".");
    }
    Serial.println();
  }

  if (WiFi.status() != WL_CONNECTED) {
    Serial.println("Failed to reconnect WiFi.");
    showLcdMessage("WiFi failed", "Check network");
    return false;
  }

  HTTPClient http;
  String url = buildScanUrl();
  String body = String("{") +
                "\"tag\":\"" + tag + "\"," +
                "\"action\":\"" + action + "\"," +
                "\"source\":\"esp32-rc522\"," +
                "\"scannedAt\":\"" + getIsoTimestamp() + "\"" +
                "}";

  secureClient.setInsecure();
  http.begin(secureClient, url);
  http.addHeader("Content-Type", "application/json");

  const int httpCode = http.PUT(body);
  const String response = http.getString();
  http.end();

  Serial.print("Firebase HTTP code: ");
  Serial.println(httpCode);
  Serial.print("Firebase response: ");
  Serial.println(response);

  if (httpCode >= 200 && httpCode < 300) {
    showLcdMessage(action + " uploaded", tag.substring(0, LCD_COLUMNS));
  } else {
    showLcdMessage("Upload failed", String(httpCode));
  }

  return httpCode >= 200 && httpCode < 300;
}

bool connectWiFi() {
  Serial.print("Connecting to WiFi");
  showLcdMessage("Connecting WiFi", "Please wait...");
  WiFi.begin(WIFI_SSID, WIFI_PASSWORD);

  const unsigned long startedAt = millis();
  while (WiFi.status() != WL_CONNECTED && millis() - startedAt < WIFI_CONNECT_TIMEOUT_MS) {
    delay(500);
    Serial.print(".");
  }

  if (WiFi.status() != WL_CONNECTED) {
    Serial.println();
    Serial.println("WiFi connection failed.");
    Serial.print("SSID used: ");
    Serial.println(WIFI_SSID);
    Serial.println("Check hotspot name, password, and 2.4GHz mode.");
    showLcdMessage("WiFi failed", "Check hotspot");
    return false;
  }

  Serial.println();
  Serial.print("Connected. IP: ");
  Serial.println(WiFi.localIP());
  showLcdMessage("WiFi connected", WiFi.localIP().toString());
  return true;
}

void setup() {
  Serial.begin(115200);
  delay(500);

  Wire.begin(LCD_SDA_PIN, LCD_SCL_PIN);
  lcd.init();
  lcd.backlight();
  showLcdMessage("FMAERO RFID", "Booting...");
  pinMode(BUTTON_STOCK_IN_PIN, INPUT_PULLUP);
  pinMode(BUTTON_STOCK_OUT_PIN, INPUT_PULLUP);

  if (!connectWiFi()) {
    Serial.println("ESP32 will retry WiFi in loop.");
    return;
  }

  SPI.begin(RFID_SCK_PIN, RFID_MISO_PIN, RFID_MOSI_PIN, RFID_SS_PIN);
  mfrc522.PCD_Init();

  Serial.println("RC522 ready.");
  Serial.println("Tap RFID card/tag now.");
  updateSelectedAction("IN");
}

void loop() {
  handleButtons();

  if (WiFi.status() != WL_CONNECTED) {
    static unsigned long lastRetryAt = 0;
    if (millis() - lastRetryAt >= 5000) {
      lastRetryAt = millis();
      connectWiFi();
      if (WiFi.status() == WL_CONNECTED) {
        SPI.begin(RFID_SCK_PIN, RFID_MISO_PIN, RFID_MOSI_PIN, RFID_SS_PIN);
        mfrc522.PCD_Init();
        Serial.println("RC522 ready.");
        Serial.println("Tap RFID card/tag now.");
        updateSelectedAction(selectedAction);
      }
    }
    delay(100);
    return;
  }

  if (!mfrc522.PICC_IsNewCardPresent()) {
    delay(50);
    return;
  }

  if (!mfrc522.PICC_ReadCardSerial()) {
    delay(50);
    return;
  }

  const String tag = readTagUid();
  const unsigned long now = millis();

  if (tag == lastTag && now - lastSentAt < DUPLICATE_BLOCK_MS) {
    Serial.print("Duplicate tag ignored: ");
    Serial.println(tag);
    showLcdMessage("Duplicate tag", tag.substring(0, LCD_COLUMNS));
  } else {
    Serial.print("Detected tag: ");
    Serial.println(tag);
    showLcdMessage("Tag " + selectedAction, tag.substring(0, LCD_COLUMNS));

    if (sendTagToFirebase(tag, selectedAction)) {
      lastTag = tag;
      lastSentAt = now;
      Serial.println("Scan pushed to Firebase.");
      delay(1000);
      updateSelectedAction(selectedAction);
    } else {
      Serial.println("Failed to push scan to Firebase.");
      delay(1200);
      updateSelectedAction(selectedAction);
    }
  }

  mfrc522.PICC_HaltA();
  mfrc522.PCD_StopCrypto1();
  delay(200);
}
