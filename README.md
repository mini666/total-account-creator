# total-account-creator
대차대조표 엑셀파일을 읽어 총계정 엑셀 파일 생성

## 인자 설명
```
usage: TotalAccountCreator [-a <arg>] [-b <arg>] [-c <arg>] [-d <arg>] [-f
       <arg>] [-h <arg>] -i <arg> -o <arg> [-s <arg>]
 -a,--targetAccounts <arg>   Processing account separated by space.
                             ex) 현금 카드
 -b,--beginRow <arg>         Start row of content.
 -c,--creditColumn <arg>     Credit column index.
 -d,--debitColumn <arg>      Debit column index.
 -f,--dateFormat <arg>       Date format for use.
 -h,--headerRow <arg>        row of header.
 -i,--input <arg>            Input file for processing.
 -o,--output <arg>           Output file.
 -s,--sheet <arg>            Sheet index.
 ```

## 실행
[apache-maven](https://dlcdn.apache.org/maven/maven-3/3.8.6/binaries/apache-maven-3.8.6-bin.tar.gz)이 설치되어 있어야 한다.  
최초 compile 실행

```
mvn compile   // 최초 한번만 실행
mvn exec:java \
-Dexec.mainClass=com.dreamer.app.TotalAccountCreator \
-Dexec.args=" \
-i <input-file-path> \
-o <onput-file-path> \
"
```
