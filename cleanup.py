"""Run from an external scheduler with DATABASE_URL set."""
import storage
if __name__=='__main__':
    storage.initialize()
    print('Expired rows removed:',storage.cleanup())
