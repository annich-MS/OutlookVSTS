import { action, computed, observable} from "mobx";
import IVSTSConfig from "../models/vstsConfig";

export default class VSTSConfigStore {

    @computed get configs(): IVSTSConfig[] { return this._configs; };
    @computed get selected(): string {return this._selected; };

    @observable private _configs: IVSTSConfig[] = [];
    @observable private _selected: string = "";

    @action public addConfig(config: IVSTSConfig): void {
        this._configs.push(config);
    }

    @action public removeConfig(configName: string): void {
        this._configs = this._configs.filter((config: IVSTSConfig) => { return config.name === configName; });
    }

    @action public setSelected(configName: string): void {
        this._selected = configName;
    }
}
