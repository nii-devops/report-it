from app import create_app
from config import DevelopmentConfig, ProductionConfig


#app = create_app(DevelopmentConfig)
app = create_app(ProductionConfig)

if __name__ == '__main__':
    print("DEBUG mode:", app.debug)
    app.run()
    
