import {ISPList} from './Demo1WebPart';
export default class MockHttpClient{
    private static _items: ISPList[] = [
        {Title:"Rajesh", Id:"1"},
        {Title:"Manish", Id:"2"},
        {Title:"Bipin", Id:"3"}
    ];

    public static get():Promise<ISPList[]>{
        return new Promise<ISPList[]>((resolve) =>
        {
            resolve (MockHttpClient._items);
        }
        );
    }
}