#include <SPI.h>
#include <nRF24L01.h>
#include <RF24.h>

int text[1];

RF24 radio(7, 8); // CE, CSN
const byte address[6] = "00001";
void setup() {
  Serial.begin(9600);
  radio.begin();
  bool check = radio.isChipConnected();
//  if(check == 1){
//    Serial.println("Chip Berhasil Connect");
//  }else{
//    Serial.println("Chip Gagal Connect");
//  }
  
  radio.openReadingPipe(0, address);
  radio.setPALevel(RF24_PA_MIN);
  radio.startListening();
}

void loop() {
  while (radio.available()) {
     radio.read(&text, sizeof(text));
     Serial.print(text[0]);
     Serial.println(',');
     
  }
  
}
