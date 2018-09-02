# Mil cardióides no Excel

Estive a folhear um livro de puzzles antigo, quando me deparo com uma curva matemática curiosa: a cardióide.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_36_2.jpg)

Esta tem este nome por parecer um coração.

Um detalhe curioso é que é possível desenhá-la somente usando régua e compasso. Como era um processo simples, mas trabalhoso para quem não tem muita coordenação motora, achei mais fácil fazer uma macro em VBA no Excel do que utilizar lápis, papel, régua e compasso.


Roteiro:
– Como desenhar a cardióide no braço
– Dicas de como usar o VBA

Link da planilha Excel para download: https://github.com/asgunzi/cardioidExcel (é necessário ativar macros para rodar).

---

# Como desenhar a cardióide no braço

O desenho da cardióide segue os seguintes passos:

Pegue uma circunferência, e a divida em n ângulos iguais, formando n “casas”.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_drawing1.jpg)


Trace uma linha num diâmetro.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_drawing2.jpg)

Trace uma linha pulando uma casa do lado direito e duas do lado esquerdo.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_drawing3.jpg)

Repita o processo n vezes.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_drawing4.jpg)

Se eu chamar de “n” o número de pontos, e de “step” o número de casas puladas, posso fazer algumas variantes desta brincadeira, com efeitos muito bonitos.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_72_2.jpg)

Em geral, quanto mais linhas, maior a resolução da cardióide – porém linhas demais tornam-o ilegível.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_33_1.jpg)

Para mudar a cor, basta colorir a célula da cor da forma desejada.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_30_0.jpg)

Quanto maior o step, mais saliências a figura vai ter.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_18_2.jpg)

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_72_3.jpg)

Embora cada curva dessas possa ter um nome, é mais fácil continuar chamando-as de cardióides.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_100_4.jpg)

E assim sucessivamente, é possível fazer cardióides a mil.

 

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_100_5.jpg)

 

 
---
#Como usar o vba do Excel para fazer este desenho

 

Da mesma forma que, no braço, é necessário apenas régua e compasso, no excel é necessário somente círculos, linhas e matemática.

 

Vou colocar apenas os pontos principais.

Para adicionar um círculo no Excel, é só usar o shape msoShapeOval.

 
```visual basic
‘Adiciona círculo na posição (left, top), com raio dado.
ActiveSheet.Shapes.AddShape(msoShapeOval, left, top, 2 * raio, 2 * raio).Select
```
 

Para traçar uma linha no Excel, utilizar o addLine do shapes, onde o ponto 1 é dado por coordenadas (x1,y1) e o ponto 2 por coordenadas (x2,y2).

 
```visual basic
‘Adiciona uma linha
ActiveSheet.Shapes.addLine(x1, y1, x2, y2).Select
```
 

Para saber quais os pontos a gerar.

Divido n pontos num círculo, o que equivale a um ângulo theta de 360 graus (ou 2*pi) dividido por n.

 

Cada ponto terá coordenadas (raio*cos(theta), raio*sin(theta)) em relação ao centro do círculo.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_angulo.jpg)

Como o centro do círculo tem coordenadas (latOrigin, longOrigin), deve-se compensar este valor. E crio um array para armazenar estas informações.

```visual basic
‘Gera coordenadas
ReDim arrRef(1 To npoints, 1 To 2) ‘lat e long

For i = 1 To npoints
  theta = (i – 1) * 2 * pi / npoints
  arrRef(i, 1) = latOrigin + raio * Cos(theta) ‘lat
  arrRef(i, 2) = longOrigin – raio * Sin(theta) ‘long
Next i
```
 
Finalmente, traço as linhas entre dois pontos (ponto i e ponto j) pulando o step dado.

```visual basic
For i = 1 To npoints

  ActiveSheet.Shapes.addLine(arrRef(i, 1), arrRef(i, 2), arrRef(j, 1), arrRef(j, 2)).Select
  ‘Update j: adiciona step
  j = (j + step – 1) Mod npoints + 1

Next i
```
 

 

A parte VBA tem que ter familiaridade com o assunto para entender, porém, os passos são exatamente os mesmos de fazer com régua e compasso, com lápis e papel!

Bônus: Cardióide com 6 saliências.

![](https://ferramentasexcelvba.files.wordpress.com/2018/09/cardiod_150_6.jpg)



--- 

Links:

[https://en.wikipedia.org/wiki/Cardioid]

[https://github.com/asgunzi/cardioidExcel]

Main blog: [https://ideiasesquecidas.com/]

Other writings: [https://medium.com/@arnaldogunzi]
