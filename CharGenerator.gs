// Função que gera uma senha secreta de 18 caracteres.
function charIdGenerator()
 {
     var charId = "";
       for (var i = 0; i < 18 ; i++) 
       { 
           charId += Math.random().toString(36); // Gera um caracter alfanumérico aleatório.
       } 
     return charId;    
 }