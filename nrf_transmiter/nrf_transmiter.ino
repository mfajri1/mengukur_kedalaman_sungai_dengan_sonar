#include <SPI.h>
#include <nRF24L01.h>
#include <RF24.h>
#include <SoftwareSerial.h>
#include <Wire.h>
#include <LiquidCrystal_I2C.h>

RF24 radio(7, 8); // CE, CSN
const byte address[6] = "00001";

int sonar = A1;
int button = 4;
int buzzer = 3;
int bacaSonar = 0;
int bacaBtn = 0;
int message[1];

LiquidCrystal_I2C lcd(0x27, 16, 2);

void setup() {  
  Serial.begin(9600);
  pinMode(sonar, INPUT);
  pinMode(button, INPUT);
  pinMode(buzzer, OUTPUT);

  lcd.begin();
  lcd.backlight();
  
  radio.begin();
  bool check = radio.isChipConnected();
  if(check == 1){
    lcd.clear();
    lcd.setCursor(0,0);
    lcd.print("Koneksi = Berhasil");
    Serial.println("Chip Berhasil Connect");
  }else{
    lcd.clear();
    lcd.setCursor(0,0);
    lcd.print("Koneksi = Gagal");
    Serial.println("Chip Gagal Connect");
  }
  
  radio.openWritingPipe(address);
  radio.setPALevel(RF24_PA_MIN);
  radio.stopListening();
  
}
void loop() {
    bacaSonar = analogRead(sonar);
    Serial.println(bacaSonar);
    digitalWrite(buzzer, HIGH);
    delay(500);
    digitalWrite(buzzer, LOW);
    lcd.clear();
    lcd.setCursor(0,0);
    lcd.print("Kedalaman");
    lcd.setCursor(0, 1);
    lcd.print(bacaSonar);
    message[0] = bacaSonar;
    radio.write(&message, sizeof(message));
    delay(1000);
}
