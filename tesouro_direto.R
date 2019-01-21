# http://www.tesouro.gov.br/tesouro-direto-balanco-e-estatisticas

library(XLConnect)
library(data.table)

# make chart labels beautiful

makelabels <- function( aux ){
  memo <- aux[1,date]
  dts <- aux[1,date]
  for(i in 2:length(aux[,date])){
    if( year( aux[i,date] ) == year( memo ) ){
      e <- NA
    }else{
      e <- aux[i,date]
      memo <- e
    } 
    dts <- c(dts,e)
  }
  dts <- format(dts, format="%Y")
  return(dts)
}

# clear empty and NA non date rows 

clearrows <- function( df ){
  clean <- data.frame()
  for(i in 1:nrow( df )){
    trig <- 0
    for(j in 2:ncol( df )){
      if( is.na(df[i,j]) | df[i,j] == '' ){
        next
      }else{
        trig <- 1
      }
    }
    if( trig == 1 ){
      clean <- rbind(clean, df[i,])
    }
  }
  return(clean)
}

# read all sheets and organize data by year and type of bond

tesouro_direto_a <- function(arq){
  ano_base <- as.numeric( gsub('^.*_|.xls','',arq) )
  print(ano_base)
  wb <- loadWorkbook(arq)
  sheets <- getSheets(wb)
  l <- list()
  # loop of sheets
  for(s in sheets){
    print(s)
    name <- gsub(' ','',s)
    df <- readWorksheet(wb, sheet = s, header = F, startRow = 3)
    # handle exception
    if( ano_base == 2008 ){
      df <- readWorksheet(wb, sheet = s, header = F, startRow = 2)
      df <- df[2:nrow(df),]
    }
    # clear empty and NA rows
    df <- clearrows( df )
    dt <- data.table(df[,c(1:3)])
    # handle exceptions
    #if( arq == 'historicoLTN_2011.xls' & name == 'LTN010711' ){
    #  dt <- data.table(df[c(1:124),c(1:3)])
    #}
    names(dt) <- c('date','bid','ask')
    #
    # organize dates
    #
    # there's a bug in the excel sheets
    # line 164 in LTN010117, LTN010118 and LTN010121 and line 148 in LTN010119 and LTN010123
    if( ano_base == 2016 ){
      dt[,date := ifelse( date == '2016-08-23 00:00:00', '23/08/2016', date )]
    }
    if( ano_base == 2008 ){
      dt[,date := as.Date(date)]
    }
    if( class( dt[1,date] ) == 'character' ){
        dt[,date := as.Date(date,"%d/%m/%Y")]
    }else{
        dt[,date := as.Date(date)]
    }
    # numbers
    #if( (ano_base == 2011 & name == 'LTN010711') & (ano_base == 2010 & name == 'LTN010710') ){
    #  dt[,bid := as.numeric(gsub('%','',bid))]
    #  dt[,ask := as.numeric(gsub('%','',ask))]
    #}else{
    #  dt[,bid := 100*as.numeric(gsub('%','',bid))]
    #  dt[,ask := 100*as.numeric(gsub('%','',ask))]
    #}
    dt[,bid := as.numeric(gsub('%','',bid))]
    dt[,ask := as.numeric(gsub('%','',ask))]
    if( dt[1,bid < 1] ){
      dt[,bid := 100*bid]
    }
    if( dt[1,ask < 1] ){
      dt[,ask := 100*ask]
    }
    # buy & sell average
    dt[,avr := (bid + ask)/2]
    tso <- dt[,list(date,avr)]
    
    # harmonize sheet names
    
    if( grepl( 'NTNBP[0-9]', name ) ){
      name <- gsub( 'BP', '-BPrincipal', name )
    }
    if( grepl( 'NTNB[0-9]', name ) ){
      name <- gsub( 'B', '-B', name )
    }
    if( grepl( 'NTNC[0-9]', name ) ){
      name <- gsub( 'C', '-C', name )
    }
    if( grepl( 'NTNF[0-9]', name ) ){
      name <- gsub( 'F', '-F', name )
    }
    # name sheet and name list object
    names(tso)[2] <- name
    #
    l[[ name ]] <- tso
  }
  return(l)
}

# monta lista com info completa sobre titulos e datas

periodo <- c(2002:2019)
l <- list()
for(ano_base in periodo){
  titulos <- c('LTN','NTN-B_Principal','NTN-B','NTN-C','NTN-F')
  for(titulo in titulos){
    # 
    if( ano_base <= 2004 & titulo == 'NTN-B_Principal' ){next}
    if( ano_base <= 2003 & titulo == 'NTN-F' ){next}
    if( ano_base == 2002 & titulo == 'NTN-B' ){next}
    # construct file name
    if( ano_base >= 2012 ){
      xls <- paste('xls/',titulo,'_',ano_base,'.xls',sep='')
    }else{
      xls <- paste('xls/historico',gsub('_|-','',titulo),'_',ano_base,'.xls',sep='')
    }
    e <- tesouro_direto_a(xls)
    for(n in names(e)){
      # organize
      if( n %in% names(l) ){
        l[[ n ]] <- rbind( l[[ n ]], e[[ n ]] )
      }else{
        l[[ n ]] <- e[[ n ]]
      }
    }
  }
}

# organize all durations in the same data table to make chart

alldur <- function( kw ){
  # loop each instrument & duration
  i <- 1
  for( n in names(l) ){
    # search for a keyword
    if( grepl( kw, n )){
      print(n)
      e <- l[[ n ]]
      setkey(e,date)
      if( i == 1 ){
        aux <- e
      }else{
        aux <- merge( aux, e, all=T )
      }
      print(aux)
      i <- i + 1
    }
  }
  return(aux)
}

#
# charts
#

# LTN: fixed rate

ltn <- alldur('LTN')
write.table(ltn,file='LTN.tsv',sep='\t',row.names = F)
lbsy <- makelabels(ltn)
ltn[,date := NULL]
matplot(ltn, type = 'l',bty='n',lty = 1,main='LTN', xaxt='n',ylab='%')
axis(side=1, at=seq(1,by=1,length.out=nrow(ltn)), labels = lbsy, tick = F, las=2, cex.axis=.7 )
abline(h=5)

# NTN-BPrinc: IPCA sem juros

ntnbp <- alldur('NTN-BPrinc')
write.table(ntnbp,file='NTN-BPrinc.tsv',sep='\t',row.names = F)
lbsy <- makelabels(ntnbp)
ntnbp[,date := NULL]
# remove bad behaved series from chart
ntnbp[,'NTN-BPrincipal150515' := NULL]
names(ntnbp)
matplot(ntnbp, type = 'l',bty='n',lty = 1,main='NTN-BPrinc', xaxt='n')
axis(side=1, at=seq(1,by=1,length.out=nrow(ntnbp)), labels = lbsy, tick = F, las=2, cex.axis=.7 )
abline(h=2)

# NTN-B: IPCA com juros semestrais

ntnb <- alldur('NTN-B[0-1]')
write.table(ntnb,file='NTN-B.tsv',sep='\t',row.names = F)
lbsy <- makelabels(ntnb)
ntnb[,date := NULL]
ntnb[,'NTN-B150515' := NULL]
ntnb[,"NTN-B150513" := NULL]
ntnb[,"NTN-B150511" := NULL]
matplot(ntnb, type = 'l',bty='n',lty = 1,main='NTN-B', xaxt='n',ylim=c(2,13))
axis(side=1, at=seq(1,by=1,length.out=nrow(ntnb)), labels = lbsy, tick = F, las=2, cex.axis=.7 )
abline(h=2)

# NTN-C: SELIC

ntnc <- alldur('NTN-C')
write.table(ntnc,file='NTN-C.tsv',sep='\t',row.names = F)
lbsy <- makelabels(ntnc)
ntnc[,date := NULL]
ntnc[,"NTN-C010311" := NULL]
ntnc[,"NTN-C010408" := NULL]
matplot(ntnc, type = 'l',bty='n',lty = 1,main='NTN-C', xaxt='n',ylim = c(0,18))
axis(side=1, at=seq(1,by=1,length.out=nrow(ntnc)), labels = lbsy, tick = F, las=2, cex.axis=.7 )
abline(h=0)

# NTN-F: Tesouro Prefixado com Juros Semestrais

ntnf <- alldur('NTN-F')
write.table(ntnf,file='NTN-F.tsv',sep='\t',row.names = F)
lbsy <- makelabels(ntnf)
ntnf[,date := NULL]
matplot(ntnf, type = 'l',bty='n',lty = 1,main='NTN-F', xaxt='n',ylim = c(5,max(ntnf,na.rm = T)))
axis(side=1, at=seq(1,by=1,length.out=nrow(ntnf)), labels = lbsy, tick = F, las=2, cex.axis=.7 )
abline(h=5)

