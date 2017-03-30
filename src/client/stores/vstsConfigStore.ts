import { action, computed, observable } from "mobx";
import IVSTSConfig from "../models/vstsConfig";
import RoamingSettings from "../models/roamingSettings";

export default class VSTSConfigStore {

    @computed get configs(): IVSTSConfig[] { return this._configs; };
    @computed get selected(): string { return this._selected; };

    @observable private _configs: IVSTSConfig[] = [];
    @observable private _selected: string = "";

    @action public setConfigs(configs: IVSTSConfig[]): void {
        this._configs = configs;
        this.resetSelected();
    }
    @action public addConfig(config: IVSTSConfig): void {
        this._configs.push(config);
        this.save();
    }

    @action public removeConfig(configName: string): void {
        this._configs = this._configs.filter((config: IVSTSConfig) => { return config.name === configName; });
        if (this._selected === configName) {
            this.resetSelected();
        }
        this.save();
    }

    @action public setSelected(configName: string): void {
        this._selected = configName;
    }

    @action private resetSelected() {
        if (this._configs.length > 0) {
            this._selected = this._configs[0].name;
        } else {
            this._selected = "";
        }
    }

    private async save(): Promise<void> {
        let rs: RoamingSettings = await RoamingSettings.GetInstance();
        rs.configs = this._configs;
        await rs.save();
    }
}

export const vstsConfig = new VSTSConfigStore();
