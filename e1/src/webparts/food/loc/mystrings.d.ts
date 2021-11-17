declare interface IFoodWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'FoodWebPartStrings' {
  const strings: IFoodWebPartStrings;
  export = strings;
}
