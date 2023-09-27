param([string]$Name, [int]$Age, $FavouriteAnimals)

if(-not $Name -or -not $Age -or -not $FavouriteAnimals){
    Write-Host "Please enter all paramneters"
    return
}
if($FavouriteAnimals -is [Array]){
    
    $Output = "Hi $Name, you are $Age years old and your favourite animals are: "
    foreach($animals in $FavouriteAnimals){
        $Output += "$animals "
    }
}
else{
    $Output = "Hi $Name, you are $Age years old and your favourite animal is a $FavouriteAnimals"
}
Write-Host $Output