Myvector <- c(0.88,0.79,0.78,0.62,0.60,0.58)

resposta <- kmeans(Myvector,3, algo="Lloyd")

resposta$cluster
data.frame(
  frame = Myvector,
  grupo = resposta$cluster
)

vetor <- c(3:6, 10:14, 20:28, 32:36)

resposta <- kmeans(vetor, 4, algorithm = "Lloyd")

data.frame(
  frame = vetor,
  grupo = resposta$cluster
)
