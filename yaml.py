import yaml

with open("configuration.yaml", "r") as f:
    data = yaml.loader(f)

print(data)