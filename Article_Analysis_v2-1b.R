# This is used to expedite analysis in a systematic review
# by Dr Marc-Andre Leger, Ph.D. 
# marcandre@leger.ca

require(pdftools) # reads pdf documents
require(tm) # text mining analysys
library("xlsx")
library ("openxlsx")
library("wordcloud")
library("wordcloud2")
library("RColorBrewer")

In_directory <- "documents2" # indicates the input directory to get the PDFs
Out_excel_File <- "words_analysis_CompliMgmt_final.xlsx" # This is the output file name

files <- list.files(In_directory, pattern="pdf$", full.names=TRUE, recursive=TRUE)
opinions <- lapply(files, pdf_text)
length(opinions) # how many files were loaded
Nb_pages <- lapply(opinions,length) # the length in pages of each PDF file
Matrix_pages <- matrix(unlist(Nb_pages))

# creating a PDF database
pdfdatabase <- Corpus(URISource(files),readerControl = list(reader = readPDF))
pdfdatabase <- tm_map(pdfdatabase, removePunctuation, ucp = TRUE)
opinions.tdm <- TermDocumentMatrix(pdfdatabase,control = list(removePunctuation = TRUE,
                                                              stopwords = TRUE,
                                                              tolower = TRUE,
                                                              stemming = FALSE,
                                                              removeNumbers = TRUE,
                                                              bounds = list(global = c(3,Inf))))
inspect(opinions.tdm[10:20,]) #examine 10 words at a time across documents
opinionstemmed.tdm <- TermDocumentMatrix(pdfdatabase,control = list(removePunctuation = TRUE,
                                                                    stopwords = TRUE,
                                                                    tolower = TRUE,
                                                                    stemming = TRUE,
                                                                    removeNumbers = TRUE,
                                                                    bounds = list(global = c(3,Inf))))
inspect(opinionstemmed.tdm[10:20,]) #examine 10 words at a time across documents
# this is inspired from https://data.library.virginia.edu/reading-pdf-files-into-r-for-text-mining/
ft <- findFreqTerms(opinions.tdm, lowfreq = 100, highfreq = Inf)
as.matrix(opinions.tdm[ft,]) 
ft.tdm <- as.matrix(opinions.tdm[ft,])
freq1 <- apply(ft.tdm, 1, sum)
df <- sort(apply(ft.tdm, 1, sum), decreasing = TRUE)
ft2 <- findFreqTerms(opinionstemmed.tdm, lowfreq = 100, highfreq = Inf)
as.matrix(opinionstemmed.tdm[ft2,]) 
ft2.tdm <- as.matrix(opinionstemmed.tdm[ft2,])
df2 <- sort(apply(ft2.tdm, 1, sum), decreasing = TRUE)

# This is the version with the words as they appear in the corpus
output1 <- data.frame(df)
output2 <- data.frame(ft.tdm)
# This is the stemmed version of the words from the corpus
output3 <- data.frame(df2)
output4 <- data.frame(ft2.tdm)
# This is a count of the Nb pages per document
output5 <- data.frame(Matrix_pages)

#define sheet names for each data frame
dataset_names <- list('Articles' = output1, 'Words' = output2, 'A_Stemmed' = output3, 'W_Stemmed' = output4, 'Pages' = output5)

#export each data frame to separate sheets in same Excel file
openxlsx::write.xlsx(dataset_names, file = Out_excel_File, colNames = TRUE, rowNames = TRUE, append = FALSE) 



# this is where the wordcloud is created
set.seed(1234)
wordcloud(words = ft, freq = freq1, min.freq = 10,
          max.words=200, random.order=FALSE, rot.per=0.35, 
          colors=brewer.pal(8, "Dark2"))

