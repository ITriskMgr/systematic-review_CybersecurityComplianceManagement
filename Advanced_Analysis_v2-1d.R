# This is used to expedite analysis in a systematic review
# it also includes the Quanteda analysis features
# by Dr Marc-Andre Leger, Ph.D. 
# marcandre@leger.ca

require(pdftools) # reads pdf documents
require(tm) # text mining analysys
require(quanteda)
require(quanteda.textstats)
library("data.table")

In_directory <- "documents2" # indicates the input directory to get the PDFs
Out_excel_File <- "Quanteda_analysis_CompliMgmt_007.xlsx" # This is the output file name for the word counts

# read the files that are used
files <- list.files(In_directory, pattern="pdf$", full.names=TRUE, recursive=TRUE)
pdfdatabase <- Corpus(URISource(files),readerControl = list(reader = readPDF))
pdfcorpus <- corpus(pdfdatabase)

# remove stopwords
my_stopwords <- read.csv("add_words.csv", header = FALSE) # add my own list of words
my_stopwords <- as.character(my_stopwords$V1)
sw <- c(my_stopwords, stopwords("en"))

#tokeize
tok_files <- tokens(pdfcorpus, 
       what = c("word"), 
       remove_numbers = TRUE, 
       remove_punct = TRUE,
       remove_symbols = TRUE, 
       remove_separators = TRUE,
       remove_url = TRUE,
       verbose = quanteda_options("verbose"), 
       include_docvars = TRUE )
tok_files <- tokens_remove(tok_files, sw)
tok_files <- tokens_tolower(tok_files)
tok_files <- tokens_wordstem(tok_files)
# tok_files

dtm <- dfm(tok_files)
doc_freq <- docfreq(dtm)
dtm <- dtm[,doc_freq >= 2]
dtm <- dfm_weight(dtm, "count")
# head(dtm)
# topfeatures(dtm)

top_bigrams <- tokens_select(tok_files, 
                             pattern = "^[A-Z]", 
                             valuetype = "regex", 
                             case_insensitive = TRUE, 
                             padding = TRUE) %>% 
  textstat_collocations(min_count = 300, size = 2, tolower = TRUE)
# top_bigrams

doc_bigrams <- tokens_select(tok_files, 
                                pattern = "^[A-Z]", 
                                valuetype = "regex", 
                                case_insensitive = TRUE, 
                                padding = TRUE) %>% 
  textstat_collocations(min_count = 5, size = 2, tolower = TRUE)

top_trigrams <- tokens_select(tok_files, 
                              pattern = "^[A-Z]", 
                              valuetype = "regex", 
                              case_insensitive = TRUE, 
                              padding = TRUE) %>% 
  textstat_collocations(min_count = 100, size = 3, tolower = TRUE)
# top_trigrams

doc_trigrams <- tokens_select(tok_files, pattern = "^[A-Z]", 
                             valuetype = "regex", 
                             case_insensitive = TRUE, 
                             padding = TRUE) %>% 
  textstat_collocations(min_count = 5, size = 3, tolower = TRUE)

top_quagrams <- tokens_select(tok_files, 
                              pattern = "^[A-Z]", 
                              valuetype = "regex", 
                              case_insensitive = TRUE, 
                              padding = TRUE) %>% 
  textstat_collocations(min_count = 50, size = 4, tolower = TRUE)
# top_quagrams

doc_quagrams <- tokens_select(tok_files, 
                              pattern = "^[A-Z]", 
                              valuetype = "regex", 
                              case_insensitive = TRUE, 
                              padding = TRUE) %>% 
  textstat_collocations(min_count = 5, size = 4, tolower = TRUE)
# head(doc_bigrams, 20)
# head(doc_trigrams, 20)

# row number too from which the words were considered
# this needs more work
# from https://stackoverflow.com/questions/65699579/table-of-n-grams-and-identifying-the-row-in-which-the-text-appeared
quadgrs <- textstat_frequency(dfm(tok_files), groups = docnames(tok_files)) %>%
  as.data.table()
setnames(quadgrs, "group", "rownumber")

row_from <- quadgrs[, c("feature", "frequency", "rownumber")]


# This is the version with the words as they appear in the corpus
output1 <- data.frame(doc_bigrams)
output2 <- data.frame(top_bigrams)
output3 <- data.frame(doc_trigrams)
output4 <- data.frame(top_trigrams)
output5 <- data.frame(doc_quagrams)
output6 <- data.frame(top_quagrams)
output7 <- convert(dtm, to = "data.frame")
output8 <- data.frame(row_from)

#define sheet names for each data frame
dataset_names <- list('Bigrams' = output1, 'TopBigrams' = output2, 'Trigrams' = output3, 'TopTrigrams' = output4, 'Quagrams' = output5, 'TopQuagrams' = output6, 'DTM' = output7, 'Rows' = output8)

#export each data frame to separate sheets in same Excel file
openxlsx::write.xlsx(dataset_names, file = Out_excel_File, colNames = TRUE, rowNames = TRUE, append = FALSE) 

