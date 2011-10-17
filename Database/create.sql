SET TERMOUT ONPROMPT Building Project tables.  Please wait.SET TERMOUT OFF

create table users
(
  username varchar2(20) unique not null,
  password varchar2(15) not null,
  name varchar2(30) not null,
  age number(3) not null,
  phone number(15),
  Email varchar2(50) unique not null,
  ip varchar2(15)  not null,
  pcname varchar2(30) unique not null
);


  