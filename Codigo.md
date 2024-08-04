CRIAÇÃO DA FICHA TÉCNICA DOS PRODUTOS DA SINTO DO BRASIL. 

Objetivo: Pela necessidade da empresa solicitou em mudar o layout da ficha técnica, propôs substituir a ferramenta Word pelo Excel 
para deixar o processo automático e fácil.

1. Foi criado o layout da ficha técnica:

![image](https://github.com/user-attachments/assets/5b763c31-31ef-4f04-b587-fc9ca11a9309)

2. Em uma segunda planilha foi criado todo o banco de dados com todos os dados e especeficações dos materiais.

3. Por meio do Visual Basic, foi criado o layout do formulario de "Emissão de ficha Técnica".

![image](https://github.com/user-attachments/assets/399c02e6-ebcd-4644-845f-b88cecd28d9c)

4. Criado o layout do formulário "Cadastrar novos produtos"

![image](https://github.com/user-attachments/assets/aa05d216-21fd-4edf-985b-f8862ddc6bce)

5. Criação dos códigos para o formulario "Emissão de ficha técnica"

•	Códigos para ser carregado ao iniciar o formulario. Utilizado para adicionar os dados do idioma e dos 
materiais da Sinto dentro do ComboBox.

```
Private Sub UserForm_Initialize()

Dim Lin As Integer

Lin = 3

    Do Until PlanBancoDeDados.Cells(Lin, 58) = ""
    
         CmbIdioma.AddItem PlanBancoDeDados.Cells(Lin, 58)
         
         Lin = Lin + 1
     
     Loop

Lin = 3
     
     Do Until PlanBancoDeDados.Cells(Lin, 62) = ""
     
     CmbMaterial.AddItem PlanBancoDeDados.Cells(Lin, 62)
     
        Lin = Lin + 1
     
     Loop

End Sub
```
•	
