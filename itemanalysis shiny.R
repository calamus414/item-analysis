library(shiny)
library(openxlsx)
library(knitr)
library(DT)

ui <- fluidPage(
  
  sidebarPanel(fileInput("file1", "匡拒郎", accept = ".xlsx"),
               numericInput("point", "–肈だ计", 5, min = 1, max = 100),
               actionButton("writedown", "块"),
               ),
  mainPanel(
    tabsetPanel(
      tabPanel("correcting",uiOutput("explanationcor"), div(DTOutput('rightornot'),style = "font-size:120%")),
      tabPanel("Summary",div(tableOutput("summaryTable"),style = "font-size:120%"),plotOutput("histlegend", height="250px"),plotOutput("scoreHistogram",width = "600px")),
      tabPanel("Item Analysis", div(tableOutput("itemTable"),style = "font-size:120%"))
      )
  )
)

server <- function(input, output) {
  output$itemTable <- renderTable({
    req(input$file1)
    test <- read.xlsx(input$file1$datapath, sheet = 1)
    eachpoint<-input$point
    correct<-test$correct
    questnumber<-length(test[,1])
    
    point<-matrix(nrow = 2,ncol = length(test[1,])-1)
    point[1,]<-colnames(test)[-1]
    
    for (i in 1:(length(test[1,])-1)) {
      point[2,i]<-sum(test[,i+1]==correct)*eachpoint
    }
    
    orderpoint<-t(point)[order(t(point)[,2]),]
    
    studentnumber<-length(test[1,])-1
    half<-as.numeric(studentnumber%/%2)
    low<-orderpoint[1:half,]
    high<-orderpoint[-(1:half),]
    
    colnames(test[,-1])
    number<-as.numeric()
    for (k in 1:length(low[,1])) {
      number[k]<-which(colnames(test)==low[k,1])
    }
    which(colnames(test)==low[41,1])
    
    highnumber<-as.numeric()
    for (k in 1:length(high[,1])) {
      highnumber[k]<-which(colnames(test)==high[k,1])
    }
    which(colnames(test)==high[41,1])
    
    
    
    
    lowori<-test[,number]
    highori<-test[,highnumber]
    Difficulty<-as.numeric()
    for (i in 1:questnumber) {
      Difficulty[i]<-round(sum(test[i,-1]==test[i,1])/studentnumber,digits = 2)
    }
    sum(test[1,-1]==test[1,1])
    
    Difficultyhigh<-as.numeric()
    for (i in 1:questnumber) {
      Difficultyhigh[i]<-sum(highori[i,]==test[i,1])/length(highnumber)
    }
    
    Difficultylow<-as.numeric()
    for (i in 1:questnumber) {
      Difficultylow[i]<-sum(lowori[i,]==test[i,1])/length(number)
    }
    
    
    Discrimination<-round(Difficultyhigh-Difficultylow,digits = 2)
    itemTable<-rbind(Difficulty,Discrimination)
    colnames(itemTable)<-1:questnumber
    rownames(itemTable)<-c('Difficulty','Discrimination')
    itemTable<-t(itemTable)
    
    studentanswer<-test[,-1]
    a<-as.numeric()
    b<-as.numeric()
    c<-as.numeric()
    d<-as.numeric()
    for (z in 1:questnumber) {
      a[z]<-round(sum(studentanswer[z,]=='a')/studentnumber,digits = 2)
      b[z]<-round(sum(studentanswer[z,]=='b')/studentnumber,digits = 2)
      c[z]<-round(sum(studentanswer[z,]=='c')/studentnumber,digits = 2)
      d[z]<-round(sum(studentanswer[z,]=='d')/studentnumber,digits = 2)
    }
    itemTable<-cbind(itemTable,correct,a,b,c,d)
    itemTable
    }
    , rownames = TRUE)
  
  output$summaryTable<-renderTable({
    req(input$file1)
    test <- read.xlsx(input$file1$datapath, sheet = 1)
    eachpoint<-input$point
    correct<-test$correct
    questnumber<-length(test[,1])
    
    point<-matrix(nrow = 2,ncol = length(test[1,])-1)
    point[1,]<-colnames(test)[-1]
    
    for (i in 1:(length(test[1,])-1)) {
      point[2,i]<-sum(test[,i+1]==correct)*eachpoint
    }
    studentnumber<-length(test[1,])-1
    キА<-round(mean(as.numeric(point[2,])),digits = 2)
    夹非畉<-round(sd(as.numeric(point[2,])),digits = 2)
    percent88<-round(quantile(as.numeric(point[2,]),0.88),digits = 2)
    percent75<-round(quantile(as.numeric(point[2,]),0.75),digits = 2)
    い计<-round(quantile(as.numeric(point[2,]),0.5),digits = 2)
    percent25<-round(quantile(as.numeric(point[2,]),0.25),digits = 2)
    percent12<-round(quantile(as.numeric(point[2,]),0.12),digits = 2)

    summaryTable<-rbind(キА,夹非畉,percent88,percent75,い计,percent25,percent12)
    t(summaryTable)
  }, rownames = FALSE, colnames = TRUE)
  
  output$rightornot<-renderDT({
    req(input$file1)
    test <- read.xlsx(input$file1$datapath, sheet = 1)
    eachpoint<-input$point
    correct<-test$correct
    questnumber<-length(test[,1])
    color_text <- function(value1, value2) {
      if (value1 == value2) {
        return(paste0('<span style="color:black">', value1, '</span>'))
      } else {
        return(paste0('<span style="color:red">', value1, '</span>'))
      }
    }
    studentanswer<-test[,-1]
    coloranswer<-studentanswer
    for (k in 1:length(studentanswer[1,])) {
      coloranswer[,k]<- mapply(color_text, studentanswer[,k], correct)
    }
    point<-matrix(nrow = 2,ncol = length(test[1,])-1)
    point[1,]<-colnames(test)[-1]
    for (i in 1:(length(test[1,])-1)) {
      point[2,i]<-sum(test[,i+1]==correct)*eachpoint
    }
    totalpoint<-point[2,]
    finaltable<-rbind(coloranswer,totalpoint)
    rownames(finaltable)<-c(1:questnumber,'total point')
    datatable(finaltable, escape = FALSE, options = list(pageLength = 20, searching = FALSE))
  }
  )
  
  output$explanationcor<-renderUI({
    req(input$file1)
    texthere<-'Red texts mean the answers are wrong, and the red texts are the wrong answers written by the students.'
    HTML(paste0('<div style="font-size: 25px;">', texthere, '</div>'))
    
    })
  
  output$histlegend<-renderPlot({
    req(input$file1)
    plot.new()
    par(mar=c(0,0,0,0))
    legend('topleft',legend = c('mean','','median','','percentage'),lwd = 4,cex = 1.5,col = c('red','white','blue','white','black'),bty = 'n')
  })
  
  output$scoreHistogram<-renderPlot({
    req(input$file1)
    test <- read.xlsx(input$file1$datapath, sheet = 1)
    eachpoint<-input$point
    correct<-test$correct
    questnumber<-length(test[,1])
    
    point<-matrix(nrow = 2,ncol = length(test[1,])-1)
    point[1,]<-colnames(test)[-1]
    
    for (i in 1:(length(test[1,])-1)) {
      point[2,i]<-sum(test[,i+1]==correct)*eachpoint
    }
    par(mar=c(5,5,0,1))
    hist(as.numeric(point[2,]),main = '',xlab = 'Point',cex.lab=2,cex.axis=2)
    abline(v=mean(as.numeric(point[2,])),lwd=2,col='red')
    abline(v=quantile(as.numeric(point[2,]),0.88),lwd=2)
    abline(v=quantile(as.numeric(point[2,]),0.75),lwd=2)
    abline(v=quantile(as.numeric(point[2,]),0.5),lwd=2,col='blue')
    abline(v=quantile(as.numeric(point[2,]),0.25),lwd=2)
    abline(v=quantile(as.numeric(point[2,]),0.12),lwd=2)
  })
  
  observeEvent(input$writedown,{
    req(input$file1)
    test <- read.xlsx(input$file1$datapath, sheet = 1)
    eachpoint<-input$point
    correct<-test$correct
    questnumber<-length(test[,1])
    
    point<-matrix(nrow = 2,ncol = length(test[1,])-1)
    point[1,]<-colnames(test)[-1]
    
    for (i in 1:(length(test[1,])-1)) {
      point[2,i]<-sum(test[,i+1]==correct)*eachpoint
    }
    
    orderpoint<-t(point)[order(t(point)[,2]),]
    
    studentnumber<-length(test[1,])-1
    half<-as.numeric(studentnumber%/%2)
    low<-orderpoint[1:half,]
    high<-orderpoint[-(1:half),]
    
    colnames(test[,-1])
    number<-as.numeric()
    for (k in 1:length(low[,1])) {
      number[k]<-which(colnames(test)==low[k,1])
    }
    which(colnames(test)==low[41,1])
    
    highnumber<-as.numeric()
    for (k in 1:length(high[,1])) {
      highnumber[k]<-which(colnames(test)==high[k,1])
    }
    which(colnames(test)==high[41,1])
    
    
    
    
    lowori<-test[,number]
    highori<-test[,highnumber]
    Difficulty<-as.numeric()
    for (i in 1:questnumber) {
      Difficulty[i]<-round(sum(test[i,-1]==test[i,1])/studentnumber,digits = 2)
    }
    sum(test[1,-1]==test[1,1])
    
    Difficultyhigh<-as.numeric()
    for (i in 1:questnumber) {
      Difficultyhigh[i]<-sum(highori[i,]==test[i,1])/length(highnumber)
    }
    
    Difficultylow<-as.numeric()
    for (i in 1:questnumber) {
      Difficultylow[i]<-sum(lowori[i,]==test[i,1])/length(number)
    }
    
    
    Discrimination<-round(Difficultyhigh-Difficultylow,digits = 2)
    itemTable<-rbind(Difficulty,Discrimination)
    colnames(itemTable)<-1:questnumber
    rownames(itemTable)<-c('Difficulty','Discrimination')
    itemTable<-t(itemTable)
    
    studentanswer<-test[,-1]
    a<-as.numeric()
    b<-as.numeric()
    c<-as.numeric()
    d<-as.numeric()
    for (z in 1:questnumber) {
      a[z]<-round(sum(studentanswer[z,]=='a')/studentnumber,digits = 2)
      b[z]<-round(sum(studentanswer[z,]=='b')/studentnumber,digits = 2)
      c[z]<-round(sum(studentanswer[z,]=='c')/studentnumber,digits = 2)
      d[z]<-round(sum(studentanswer[z,]=='d')/studentnumber,digits = 2)
    }
    itemTable<-cbind(itemTable,correct,a,b,c,d)
    write.csv(itemTable,'itemanalysis.csv')
    
    studentnumber<-length(test[1,])-1
    キА<-round(mean(as.numeric(point[2,])),digits = 2)
    夹非畉<-round(sd(as.numeric(point[2,])),digits = 2)
    percent88<-round(quantile(as.numeric(point[2,]),0.88),digits = 2)
    percent75<-round(quantile(as.numeric(point[2,]),0.75),digits = 2)
    い计<-round(quantile(as.numeric(point[2,]),0.5),digits = 2)
    percent25<-round(quantile(as.numeric(point[2,]),0.25),digits = 2)
    percent12<-round(quantile(as.numeric(point[2,]),0.12),digits = 2)
    
    summaryTable<-rbind(キА,夹非畉,percent88,percent75,い计,percent25,percent12)
    write.csv(t(summaryTable),'summary.csv', row.names=FALSE)
    
    totalpoint<-point[2,]
    resulttable<-rbind(studentanswer,totalpoint)
    rownames(resulttable)<-c(1:questnumber,'total point')
    write.csv(resulttable,'result.csv')
    
  })
}

# Run the application 
shinyApp(ui = ui, server = server)

