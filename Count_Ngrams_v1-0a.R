# This is used to expedite analysis in a systematic review
# it counts ngrams that are related to this project
# the ngrams are in the my_project_Words file
# by Dr Marc-Andre Leger, Ph.D. 
# marcandre@leger.ca

require(pdftools) # reads pdf documents
require(tm) # text mining analysys
require(quanteda)
require(quanteda.textstats)
require(stringi)

In_directory <- "documents2" # indicates the input directory to get the PDFs
Out_excel_File <- "Ngram_analysis_CompliMgmt_001.xlsx" # This is the output file name for the word counts

# read the files that are used
files <- list.files(In_directory, pattern="pdf$", full.names=TRUE, recursive=TRUE)
pdfdatabase <- Corpus(URISource(files),readerControl = list(reader = readPDF))
pdfcorpus <- corpus(pdfdatabase)

# load the ngrams to document
my_project_Words <- read.csv("my_project_Words.csv", header = FALSE) # add my list of words fot the project
lis_my_project_Words <- as.list(my_project_Words)

# remove stopwords
my_stopwords <- read.csv("add_words.csv", header = FALSE) # add my own list of words
my_stopwords <- as.character(my_stopwords$V1)
sw <- c(my_stopwords, stopwords())

#tokeize the corpus
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
#tok_files

lis <- lapply(lis_my_project_Words, function(x) stri_c_list(as.list(tokens(x)), sep = " "))
dict <- dictionary(lis)

## tokenize texts and count
dfm_files <- dfm(tokens_lookup(tok_files, dict))

output1 <- convert(dfm_files, to = "data.frame")
# output2 <- convert(tok_files, to = "data.frame")
output2 <- convert(dfm(tok_files), to = "data.frame")

#define sheet names for each data frame
dataset_names <- list('dfm' = output1, 'tok' = output2)

#export each data frame to separate sheets in same Excel file
openxlsx::write.xlsx(dataset_names, file = Out_excel_File, colNames = TRUE, rowNames = TRUE, append = FALSE) 

