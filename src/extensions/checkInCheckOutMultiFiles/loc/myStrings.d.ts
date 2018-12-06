declare interface ICheckInCheckOutMultiFilesCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'CheckInCheckOutMultiFilesCommandSetStrings' {
  const strings: ICheckInCheckOutMultiFilesCommandSetStrings;
  export = strings;
}
