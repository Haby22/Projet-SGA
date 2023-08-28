#include <SPI.h>
#include <MFRC522.h>

#define SS_PIN 10 //RX slave select
#define RST_PIN 9
String SALLE = "A11";
MFRC522 mfrc522(SS_PIN, RST_PIN); // Create MFRC522 instance.

byte card_ID[4]; //card UID size 4byte
// UID   = byte[4]
// UID[] = byte[][4]
typedef byte uid_t[4];
typedef struct {
  uid_t uid;
  char* name;
} person_t;
//
person_t students[] = {
  { { 0x53,0xAE,0x72,0x34 }, "WAZAMA" },
  { { 0x45,0x51,0x2E,0x83 }, "DOUCH" },
  { { 0x35,0x7C,0x2E,0x83 }, "Sanon" },
  { { 0xE8,0xC3,0xF7,0x0D }, "Haby" }
  
};
//
#define SUID      (sizeof(uid))
#define NSTUDENTS (sizeof(students) / sizeof(person_t))
//
unsigned int findIdentifier(byte* uid) {
  int n;
  n = NSTUDENTS;
  for (int i = 0; i < n; i++) {
    if (memcmp(uid, &students[i].uid, SUID)) continue;
    return i;
  }
  return -1;
}
//
byte Name1[4]={0x42,0x7E,0x19,0x1F};//first UID card
byte Name2[4]={0x45,0x51,0x2E,0x83};//second UID card

//if you want the arduino to detect the cards only once
int NumbCard[2];//this array content the number of cards. in my case i have just two cards.
int j=0;        

int const RedLed=6;
int const GreenLed=5;
int const Buzzer=8;

String Name;//user name
long Number;//user number
int n ;//The number of card you want to detect (optional)  

void setup() {
  Serial.begin(9600); // Initialize serial communications with the PC
  SPI.begin();  // Init SPI bus
  mfrc522.PCD_Init(); // Init MFRC522 card
  
  //Serial.println("CLEARSHEET");                 // clears starting at row 1
  //Serial.println("LABEL,Date,Time,Name,Number");// make four columns (Date,Time,[Name:"user name"]line 48 & 52,[Number:"user number"]line 49 & 53)

  pinMode(RedLed,OUTPUT);
  pinMode(GreenLed,OUTPUT);
  pinMode(Buzzer,OUTPUT);

   }
    
void loop() {
  

  int id;
  char buffer[256];
  person_t* person;
  //
  //look for new card
   if ( ! mfrc522.PICC_IsNewCardPresent()) {
  return;//got to start of loop if there is no card present
 }
 // Select one of the cards
 if ( ! mfrc522.PICC_ReadCardSerial()) {
  return;//if read card serial(0) returns 1, the uid struct contians the ID of the read card.
 }
 ///////////
 id = findIdentifier(mfrc522.uid.uidByte);
 if (id != -1) {
  // if ID exists, proceed and get
  // the student name;
  person = &students[id];
  //Serial.print(person->name);
  Serial.print(SALLE);
  Serial.print(',');
  Serial.println(id);
 }
 //if (! memcmp(Name1, mfrc522.uid.uidByte, sizeof(Name1))) Serial.println("First Employee, 123456");
 //if (! memcmp(Name2, mfrc522.uid.uidByte, sizeof(Name2))) Serial.println("Second Employee, 789101");
 ///////////
 /*
 for (byte i = 0; i < mfrc522.uid.size; i++) {
    // TTTTTTT
    memcmp(Name1, mfrc522.uid.uidByte, sizeof(Name1));
    memcmp(Name2, mfrc522.uid.uidByte, sizeof(Name2));
    //////////
     card_ID[i]=mfrc522.uid.uidByte[i];

       if(card_ID[i]==Name1[i]){
       Name="First Employee";//user name
       Number=123456;//user number
       j=0;//first number in the NumbCard array : NumbCard[j]
      }
      else if(card_ID[i]==Name2[i]){
       Name="Second Employee";//user name
       Number=789101;//user number
       j=1;//Second number in the NumbCard array : NumbCard[j]
      }
      else{
          digitalWrite(GreenLed,LOW);
          digitalWrite(RedLed,HIGH);
          goto cont;//go directly to line 85
     }
}
      if(NumbCard[j] == 1){//to check if the card already detect
      //if you want to use LCD
      //Serial.println("Already Exist");
      }
      else{
      NumbCard[j] = 1;//put 1 in the NumbCard array : NumbCard[j]={1,1} to let the arduino know if the card was detecting 
      n++;//(optional)
      //Serial.print("DATA,DATE,TIME," + Name);//send the Name to excel
      //Serial.print(",");
      //Serial.println(Number); //send the Number to excel
      Serial.println(Name + "," + Number);
      digitalWrite(GreenLed,HIGH);
      digitalWrite(RedLed,LOW);
      digitalWrite(Buzzer,HIGH);
      delay(30);
      digitalWrite(Buzzer,LOW);
      //Serial.println("SAVEWORKBOOKAS,Names/WorkNames");
      }
      */
      delay(1000);
cont:
delay(2000);
digitalWrite(GreenLed,LOW);
digitalWrite(RedLed,LOW);

//if you want to close the Excel when all card had detected and save Excel file in Names Folder. in my case i have just 2 card (optional)
//if(n==2){
    
  //  Serial.println("FORCEEXCELQUIT");
 //   }
 while (Serial.available() > 0) {
    
 printf('agzhejr');
    };
}
