void setup() {
  Serial.begin(9600);
  pinMode(13, OUTPUT);
}
void loop() {
  if (Serial.available()) {
    char a = Serial.read();
    if(a == '1') {
      digitalWrite(13, 1);
    }
    else if (a == '0'){
      digitalWrite(13, 0);
    }
    delay(100);
  }
}