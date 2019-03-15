"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var reptiles = /** @class */ (function () {
    function reptiles(animalDetail) {
        this.animalDetail = animalDetail;
    }
    reptiles.prototype.getreptile = function () {
        return this.animalDetail;
    };
    return reptiles;
}());
function sayName(animal) {
    var lezzy = new reptiles(animal);
    return lezzy;
}
exports.sayName = sayName;
// sayName({
//     animalName : 'carlos',
//     animalBread : 'pug',
//     waterBirdName : 'sulo',
//     waterBirdBread : 'penguine'
// })
