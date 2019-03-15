interface Animal {
    animalName: String;
    animalBread: String;
}
interface Bird {
    birdName: String;
    birdBread: String;
}
interface WaterBird {
    waterBirdName: String;
    waterBirdBread: String;
}
declare type creature = Animal & (Bird | WaterBird);
declare class reptiles {
    animalDetail: creature;
    constructor(animalDetail: creature);
    getreptile(): creature;
}
export declare function sayName(animal: creature): reptiles;
export {};
