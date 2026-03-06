class DBQuery {

 constructor(table){
   this.data = getTable(table);
 }

 where(field, value){
   this.data = this.data.filter(r => r[field] == value);
   return this;
 }

 get(){
   return this.data;
 }

 first(){
   return this.data[0] || null;
 }

}