# TesteAnalista
Automação em VBA com Selenium
Preechimento de planilha excel com dados extraídos da web.
////


Sub READY()
    
    Set driver = New ChromeDriver
    Dim a70 As WebElement
    Dim MotoOne As WebElement
    Dim Note7 As WebElement
    Dim valores As WebElements
    Dim primeiroElemento As WebElement
    
    
    driver.Get "https://www.amazon.com.br/Smartphone-Samsung-Galaxy-128Gb-Sm-A705Mzkjzto/dp/B07RJTWXL8"
    
    Sheets("PRONTO").Range("B2") = "https://www.amazon.com.br/Smartphone-Samsung-Galaxy-128Gb-Sm-A705Mzkjzto/dp/B07RJTWXL8"
    Set a70 = driver.FindElementById("productTitle")
    Sheets("PRONTO").Range("B3") = (a70.Text)
    a70valor = (driver.FindElementByClass("a-price-whole").Text)
    Sheets("PRONTO").Range("B4") = "R$" & a70valor
      
    
End Sub

Sub READY2()

    driver.Get "https://www.magazineluiza.com.br/smartphone-motorola-one-vision-128gb-azul-safira-4g-4gb-ram-634-cam-dupla-cam-selfie-25mp/p/hb13f6d686/te/srmt/?&seller_id=ijvarejo&utm_source=bing&utm_medium=pla&partner_id=65140&msclkid=53ba48acde9311dfd10a2a89f01928d3"
    
    Sheets("PRONTO").Range("B7") = "https://www.magazineluiza.com.br/smartphone-motorola-one-vision-128gb-azul-safira-4g-4gb-ram-634-cam-dupla-cam-selfie-25mp/p/hb13f6d686/te/srmt/?&seller_id=ijvarejo&utm_source=bing&utm_medium=pla&partner_id=65140&msclkid=53ba48acde9311dfd10a2a89f01928d3"
    Set MotoOne = driver.FindElementByCss("body > div.wrapper__main > div.wrapper__content.js-wrapper-content > div.wrapper__control > div.header-product.js-header-product > h1")
    Sheets("PRONTO").Range("B8") = (MotoOne.Text)
    MotoOnevalor = (driver.FindElementByClass("price-template__text").Text)
    Sheets("PRONTO").Range("B9") = "R$" & MotoOnevalor
    
End Sub

Sub READY3()

    driver.Get "https://www.amazon.com.br/Celular-Xiaomi-Redmi-Vers%C3%A3o-Global/dp/B07PB8TYCJ"
    
    Sheets("PRONTO").Range("B12") = "https://www.amazon.com.br/Celular-Xiaomi-Redmi-Vers%C3%A3o-Global/dp/B07PB8TYCJ"
    Set Note7 = driver.FindElementById("productTitle")
    Sheets("PRONTO").Range("B13") = (Note7.Text)
    Note7valor = (driver.FindElementByClass("a-price-whole").Text)
    Sheets("PRONTO").Range("B14") = "R$" & Note7valor
    
    
    MsgBox ("Aplicação chegou ao fim")
End Sub

Sub Together()

    Call READY
    Call READY2
    Call READY3
    

End Sub
