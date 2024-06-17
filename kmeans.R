Myvector <- c(0.88,0.79,0.78,0.62,0.60,0.58)

resposta <- kmeans(Myvector,3, algo="Lloyd")

resposta$cluster
data.frame(
  frame = Myvector,
  grupo = resposta$cluster
)

vetor <- c(3:6, 10:14, 20:28, 32:36)

resposta <- kmeans(linhas_embranco, 4, algorithm = "Lloyd")

resposta$cluster

data.frame(
  frame = vetor,
  grupo = resposta$cluster
)
