

interface Animal {
    animalName: String;
    animalBread: String;
}

interface Bird {
    birdName: String;
    birdBread: String;
}

interface WaterBird{
    waterBirdName : String;
    waterBirdBread : String;
}

interface insects {
    insectName : String;
    insectBread : String;
}

type creature =  Animal & (Bird | WaterBird)

class reptiles {

    animalDetail : creature;

    constructor(animalDetail : creature){
        this.animalDetail = animalDetail;
    }

    getreptile(){
        return this.animalDetail;
    }
}

export function sayName(animal : creature) : reptiles{
    let lezzy = new reptiles(animal);
    return lezzy;
}


// sayName({
//     animalName : 'carlos',
//     animalBread : 'pug',
//     waterBirdName : 'sulo',
//     waterBirdBread : 'penguine'
// })