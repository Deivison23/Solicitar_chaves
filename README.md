# Solicitar_chaves
 Esse código é uma aplicação web com o algoritmo escrito em python utilizando algumas bibliotecas em especial a streamlit para solicitação de chaves entre setores de uma companhia (empresa).
O solicitante da chave faz a solicitação preenchendo os campos obrigatórios como e-mail, datas de retirada e entrega e qual chave ele quer solicitar.
Após a solicitação ser realizada, é disparado um e-mail automático para o solicitante confirmando que a solicitação foi feita, com as informações que o solicitante preencheu e com o status em pendente, aguardando a confirmação do administrador. 
O administrador da chave, recebe um e-mail informando que uma solicitação foi aberta com as informações do solicitante no corpo do e-mail e que precisa ser liberada ou negada na plataforma. 
O administrador acessa a plataforma na aba de solicitação e só consegue acessar mediante a uma tela de login que só o administrador tem acesso.
No caso se a chave for liberada, é disparado um e-mail para o solicitante com a atualização de status aprovada, e para ser retirada com o aprovador (nome do aprovador no corpo do e-mail). OBS: A chave só é liberada ou negada com o preenchimento do nome no campo "Nome do Aprovador".
Caso a chave for negada, é disparado um e-mail para o solicitante dizendo que a chave foi negada e que no momento a chave não se encontra no setor, para efetuar a solicitação no outro dia.
Na parte de devolução da chave, o solicitante vai até o setor do administrador para a entrega da chave, o administrador recebe a chave e entra na plataforma na aba de devolução, mediante ao seu login de acesso e as informações daquela solicitação (também com um ID para fazer as contagens das solicitações) preenche no campo nome e confirma a devolução.
Conforme for efetuando as solicitações e devoluções, tem um arquivo em xlsx(excel) que é atualizado com as informações e criando um banco de dados para caso queira efetuar uma consulta.
